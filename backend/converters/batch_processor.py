import os
import gc
import time
import logging
from abc import ABC, abstractmethod
from typing import List, Dict, Any, Callable, Optional, Tuple
from dataclasses import dataclass, field
from enum import Enum
from concurrent.futures import ThreadPoolExecutor, Future, as_completed
from concurrent.futures._base import CANCELLED
import psutil
import threading

logger = logging.getLogger(__name__)


class ProcessingStatus(Enum):
    PENDING = "pending"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"
    SKIPPED = "skipped"
    CANCELLED = "cancelled"


@dataclass
class FileTask:
    file_path: str
    output_path: Optional[str] = None
    status: ProcessingStatus = ProcessingStatus.PENDING
    error_message: Optional[str] = None
    start_time: Optional[float] = None
    end_time: Optional[float] = None
    retry_count: int = 0
    file_size: int = 0
    estimated_memory: int = 0

    @property
    def duration(self) -> Optional[float]:
        if self.start_time and self.end_time:
            return self.end_time - self.start_time
        return None


@dataclass
class ProgressInfo:
    current_file: str
    file_index: int
    total_files: int
    progress_percent: float
    completed_files: int
    failed_files: int
    estimated_remaining_time: Optional[float]
    current_memory_usage: float
    is_cancelled: bool = False


@dataclass
class BatchResult:
    total_files: int
    success_count: int
    failed_count: int
    skipped_count: int
    total_duration: float
    tasks: List[FileTask] = field(default_factory=list)
    error_report: Dict[str, str] = field(default_factory=dict)
    memory_stats: Dict[str, float] = field(default_factory=dict)


class ProgressCallback:
    def __init__(self, callback_fn: Callable[[ProgressInfo], None]):
        self.callback_fn = callback_fn
        self._cancelled = False
        self._lock = threading.Lock()

    def report(self, info: ProgressInfo):
        with self._lock:
            if self._cancelled:
                info.is_cancelled = True
        self.callback_fn(info)

    def cancel(self):
        with self._lock:
            self._cancelled = True

    def is_cancelled(self) -> bool:
        with self._lock:
            return self._cancelled


class MemoryMonitor:
    def __init__(self, threshold_percent: float = 70.0):
        self.threshold_percent = threshold_percent
        self.process = psutil.Process(os.getpid())
        self.peak_memory = 0
        self._high_memory_count = 0
        self._lock = threading.Lock()

    def get_current_memory_mb(self) -> float:
        memory_info = self.process.memory_info()
        current_mb = memory_info.rss / (1024 * 1024)
        with self._lock:
            if current_mb > self.peak_memory:
                self.peak_memory = current_mb
        return current_mb

    def get_memory_percent(self) -> float:
        return self.process.memory_percent()

    def is_memory_high(self) -> bool:
        return self.get_memory_percent() > self.threshold_percent

    def should_degrade_to_serial(self) -> bool:
        with self._lock:
            if self.is_memory_high():
                self._high_memory_count += 1
            else:
                self._high_memory_count = 0
            return self._high_memory_count >= 3

    def estimate_file_memory_mb(self, file_path: str) -> float:
        if not os.path.exists(file_path):
            return 0
        file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
        return file_size_mb * 2.5

    def get_stats(self) -> Dict[str, float]:
        return {
            'current_memory_mb': self.get_current_memory_mb(),
            'peak_memory_mb': self.peak_memory,
            'memory_percent': self.get_memory_percent(),
            'threshold_percent': self.threshold_percent
        }


class FileConverter(ABC):
    @abstractmethod
    def convert(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        pass

    @abstractmethod
    def cleanup(self):
        pass


class BatchProcessor:
    def __init__(
        self,
        max_workers: Optional[int] = None,
        enable_parallel: bool = True,
        memory_threshold: float = 70.0,
        max_retries: int = 2,
        page_flush_interval: int = 50
    ):
        if max_workers is None:
            max_workers = max(1, os.cpu_count() // 2)
        
        self.max_workers = max_workers
        self.enable_parallel = enable_parallel and max_workers > 1
        self.memory_threshold = memory_threshold
        self.max_retries = max_retries
        self.page_flush_interval = page_flush_interval
        
        self.memory_monitor = MemoryMonitor(threshold_percent=memory_threshold)
        self.logger = logging.getLogger(self.__class__.__name__)
        self._current_tasks: Dict[str, FileTask] = {}
        self._lock = threading.Lock()
        self._cancelled = False

    def process_files(
        self,
        files: List[str],
        converter: FileConverter,
        output_dir: str,
        progress_callback: Optional[ProgressCallback] = None,
        extensions: Optional[List[str]] = None
    ) -> BatchResult:
        if not files:
            return BatchResult(
                total_files=0,
                success_count=0,
                failed_count=0,
                skipped_count=0,
                total_duration=0.0
            )

        os.makedirs(output_dir, exist_ok=True)
        
        tasks = self._prepare_tasks(files, extensions)
        total_duration = time.time()
        start_time = time.time()

        for task in tasks:
            task.file_size = os.path.getsize(task.file_path) if os.path.exists(task.file_path) else 0
            task.estimated_memory = self.memory_monitor.estimate_file_memory_mb(task.file_path)

        actual_workers = self._determine_worker_count(tasks)

        if actual_workers > 1 and self.enable_parallel:
            self.logger.info(f"Using parallel processing with {actual_workers} workers")
            success_count, failed_count, skipped_count, tasks = self._process_parallel(
                tasks, converter, output_dir, progress_callback
            )
        else:
            self.logger.info("Using serial processing")
            success_count, failed_count, skipped_count, tasks = self._process_serial(
                tasks, converter, output_dir, progress_callback
            )

        total_duration = time.time() - start_time

        result = BatchResult(
            total_files=len(files),
            success_count=success_count,
            failed_count=failed_count,
            skipped_count=skipped_count,
            total_duration=total_duration,
            tasks=tasks,
            error_report=self._generate_error_report(tasks),
            memory_stats=self.memory_monitor.get_stats()
        )

        self._log_summary(result)
        return result

    def _determine_worker_count(self, tasks: List[FileTask]) -> int:
        if not self.enable_parallel:
            return 1
        
        large_file_count = sum(1 for task in tasks if task.file_size > 100 * 1024 * 1024)
        
        if large_file_count > len(tasks) * 0.3:
            return 1
        
        if self.memory_monitor.should_degrade_to_serial():
            self.logger.warning("Memory threshold exceeded, degrading to serial processing")
            return 1
        
        return min(self.max_workers, max(1, os.cpu_count() // 2))

    def _prepare_tasks(self, files: List[str], extensions: Optional[List[str]]) -> List[FileTask]:
        tasks = []
        for file_path in files:
            if os.path.exists(file_path) and os.path.isfile(file_path):
                if extensions:
                    _, ext = os.path.splitext(file_path.lower())
                    if ext not in extensions:
                        continue
                tasks.append(FileTask(file_path=file_path))
        return tasks

    def _process_parallel(
        self,
        tasks: List[FileTask],
        converter: FileConverter,
        output_dir: str,
        progress_callback: Optional[ProgressCallback]
    ) -> Tuple[int, int, int, List[FileTask]]:
        success_count = 0
        failed_count = 0
        skipped_count = 0
        completed = 0
        total = len(tasks)
        start_time = time.time()
        tasks_lock = threading.Lock()

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            future_to_task = {}
            
            for task in tasks:
                output_path = self._generate_output_path(task.file_path, output_dir)
                future = executor.submit(
                    self._convert_single_file,
                    task, converter, output_path, tasks_lock
                )
                future_to_task[future] = task

            for future in as_completed(future_to_task):
                if self._cancelled:
                    executor.shutdown(wait=False, cancel_futures=True)
                    break

                task = future_to_task[future]
                try:
                    success, error_msg = future.result()
                    
                    with tasks_lock:
                        if success:
                            task.status = ProcessingStatus.COMPLETED
                            success_count += 1
                        elif task.status == ProcessingStatus.SKIPPED:
                            skipped_count += 1
                        else:
                            task.status = ProcessingStatus.FAILED
                            task.error_message = error_msg
                            failed_count += 1
                        
                        completed += 1
                        elapsed = time.time() - start_time
                        avg_time = elapsed / completed if completed > 0 else 0
                        remaining = avg_time * (total - completed)
                        
                        self._report_progress(
                            progress_callback, task.file_path, completed, total,
                            remaining, task.status == ProcessingStatus.COMPLETED
                        )
                        
                        gc.collect()
                        
                except Exception as e:
                    with tasks_lock:
                        task.status = ProcessingStatus.FAILED
                        task.error_message = str(e)
                        failed_count += 1
                        completed += 1

        return success_count, failed_count, skipped_count, tasks

    def _process_serial(
        self,
        tasks: List[FileTask],
        converter: FileConverter,
        output_dir: str,
        progress_callback: Optional[ProgressCallback]
    ) -> Tuple[int, int, int, List[FileTask]]:
        success_count = 0
        failed_count = 0
        skipped_count = 0
        start_time = time.time()
        total = len(tasks)

        for idx, task in enumerate(tasks):
            if self._cancelled:
                task.status = ProcessingStatus.CANCELLED
                break

            output_path = self._generate_output_path(task.file_path, output_dir)
            success, error_msg = self._convert_single_file(task, converter, output_path)

            if success:
                task.status = ProcessingStatus.COMPLETED
                success_count += 1
            elif task.status == ProcessingStatus.SKIPPED:
                skipped_count += 1
            else:
                task.status = ProcessingStatus.FAILED
                task.error_message = error_msg
                failed_count += 1

            elapsed = time.time() - start_time
            avg_time = elapsed / (idx + 1) if idx + 1 > 0 else 0
            remaining = avg_time * (total - idx - 1)

            self._report_progress(
                progress_callback, task.file_path, idx + 1, total, remaining,
                task.status == ProcessingStatus.COMPLETED
            )

            gc.collect()

        return success_count, failed_count, skipped_count, tasks

    def _convert_single_file(
        self,
        task: FileTask,
        converter: FileConverter,
        output_path: str,
        lock: Optional[threading.Lock] = None
    ) -> Tuple[bool, Optional[str]]:
        if self._cancelled:
            if lock:
                with lock:
                    task.status = ProcessingStatus.CANCELLED
            else:
                task.status = ProcessingStatus.CANCELLED
            return False, "Cancelled"

        task.start_time = time.time()
        
        if lock:
            with lock:
                task.status = ProcessingStatus.PROCESSING
        else:
            task.status = ProcessingStatus.PROCESSING

        try:
            success, message = converter.convert(task.file_path, output_path)
            task.end_time = time.time()
            
            if success:
                task.output_path = output_path
                return True, None
            else:
                task.status = ProcessingStatus.SKIPPED
                return False, message
                
        except Exception as e:
            task.end_time = time.time()
            error_msg = str(e)
            self.logger.error(f"Failed to convert {task.file_path}: {error_msg}")
            return False, error_msg
        finally:
            try:
                converter.cleanup()
            except Exception:
                pass

    def _report_progress(
        self,
        progress_callback: Optional[ProgressCallback],
        current_file: str,
        completed: int,
        total: int,
        remaining_time: float,
        is_success: bool
    ):
        if progress_callback:
            info = ProgressInfo(
                current_file=current_file,
                file_index=completed,
                total_files=total,
                progress_percent=(completed / total * 100) if total > 0 else 0,
                completed_files=completed,
                failed_files=0,
                estimated_remaining_time=remaining_time,
                current_memory_usage=self.memory_monitor.get_current_memory_mb(),
                is_cancelled=self._cancelled
            )
            progress_callback.report(info)

            if progress_callback.is_cancelled():
                self._cancelled = True

    def _generate_output_path(self, input_path: str, output_dir: str) -> str:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        return os.path.join(output_dir, base_name)

    def _generate_error_report(self, tasks: List[FileTask]) -> Dict[str, str]:
        error_report = {}
        for task in tasks:
            if task.status == ProcessingStatus.FAILED and task.error_message:
                error_report[task.file_path] = task.error_message
        return error_report

    def _log_summary(self, result: BatchResult):
        self.logger.info(f"Batch processing completed: {result.success_count}/{result.total_files} succeeded")
        if result.failed_count > 0:
            self.logger.warning(f"Failed files: {result.failed_count}")
        self.logger.info(f"Peak memory usage: {result.memory_stats.get('peak_memory_mb', 0):.2f} MB")

    def cancel(self):
        self._cancelled = True


class PagedFileProcessor:
    def __init__(self, page_size: int = 50):
        self.page_size = page_size
        self.current_page = 0
        self.total_pages = 0
        self._lock = threading.Lock()

    def should_flush(self, current_page: int) -> bool:
        return current_page > 0 and current_page % self.page_size == 0

    def reset(self):
        with self._lock:
            self.current_page = 0

    def increment_page(self):
        with self._lock:
            self.current_page += 1
            return self.should_flush(self.current_page)

    def set_total_pages(self, total: int):
        with self._lock:
            self.total_pages = total


class StreamingFileHandler:
    @staticmethod
    def stream_copy(src_path: str, dst_path: str, chunk_size: int = 8192) -> bool:
        try:
            with open(src_path, 'rb') as src:
                with open(dst_path, 'wb') as dst:
                    while True:
                        chunk = src.read(chunk_size)
                        if not chunk:
                            break
                        dst.write(chunk)
            return True
        except Exception as e:
            logger.error(f"Streaming copy failed: {e}")
            return False

    @staticmethod
    def safe_delete(file_path: str) -> bool:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
        except Exception as e:
            logger.error(f"Failed to delete {file_path}: {e}")
        return False
