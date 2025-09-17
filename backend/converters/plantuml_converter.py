from typing import List, Optional, Dict, Any
import os
import subprocess
import tempfile
import shutil
import logging
from pathlib import Path
from dataclasses import dataclass, field
import re
import time

from .base_converter import BaseConverter


@dataclass
class PlantUMLConfig:
    """PlantUML转换配置"""
    dpi: int = 300
    format: str = 'png'
    charset: str = 'UTF-8'
    plantuml_jar_path: Optional[str] = None
    java_options: List[str] = field(default_factory=list)
    graphviz_dot_path: Optional[str] = None
    timeout: int = 60  # 转换超时时间（秒）


@dataclass
class DependencyStatus:
    """依赖检查状态"""
    java_available: bool
    java_version: Optional[str]
    plantuml_jar_path: Optional[str]
    plantuml_version: Optional[str]
    graphviz_available: bool
    graphviz_version: Optional[str]
    
    @property
    def is_ready(self) -> bool:
        """是否准备就绪"""
        return self.java_available and bool(self.plantuml_jar_path)


class PlantUMLConverter(BaseConverter):
    """
    PlantUML到PNG转换器
    继承BaseConverter，专门处理PlantUML文件转换
    """
    
    # 支持的文件扩展名
    SUPPORTED_EXTENSIONS = ['.puml', '.plantuml', '.pu']
    
    # PlantUML JAR文件名
    PLANTUML_JAR_NAME = 'plantuml.jar'
    
    # 默认配置
    DEFAULT_CONFIG = {
        'dpi': 300,
        'format': 'png',
        'charset': 'UTF-8',
        'timeout': 60
    }
    
    def __init__(self, output_dir: str, **kwargs):
        super().__init__(output_dir, **kwargs)
        
        # 过滤PlantUMLConfig支持的参数
        plantuml_config_keys = {
            'dpi', 'format', 'charset', 'plantuml_jar_path', 
            'java_options', 'graphviz_dot_path', 'timeout'
        }
        
        # 合并默认配置，只包含PlantUMLConfig支持的参数
        config = {**self.DEFAULT_CONFIG}
        for key, value in kwargs.items():
            if key in plantuml_config_keys:
                config[key] = value
        
        self.plantuml_config = PlantUMLConfig(**config)
        
        # 依赖状态缓存
        self._dependency_status: Optional[DependencyStatus] = None
        self._jar_path_cache: Optional[str] = None
        
        # 检查依赖
        self._check_dependencies()
    
    def convert(self, input_path: str) -> List[str]:
        """
        转换PlantUML文件为PNG
        
        Args:
            input_path: 输入的PlantUML文件或包含PlantUML文件的目录
            
        Returns:
            List[str]: 生成的输出文件路径列表
        """
        # 验证输入
        if not self._is_valid_input(input_path, self.SUPPORTED_EXTENSIONS):
            raise ValueError(f"无效的输入文件或目录: {input_path}")
        
        # 检查依赖是否就绪
        if not self._dependency_status or not self._dependency_status.is_ready:
            raise RuntimeError("PlantUML转换依赖未就绪，请检查Java环境和PlantUML JAR包")
        
        output_files = []
        
        if os.path.isfile(input_path):
            # 单文件转换
            result = self._convert_single_file(input_path)
            if result:
                output_files.append(result)
        else:
            # 目录批量转换
            plantuml_files = self._get_files_by_extension(input_path, self.SUPPORTED_EXTENSIONS)
            for file_path in plantuml_files:
                try:
                    result = self._convert_single_file(file_path)
                    if result:
                        output_files.append(result)
                except Exception as e:
                    self.logger.error(f"转换文件 {file_path} 失败: {e}")
                    continue
        
        self.logger.info(f"PlantUML转换完成，共生成 {len(output_files)} 个文件")
        return output_files
    
    def _convert_single_file(self, file_path: str) -> Optional[str]:
        """
        转换单个PlantUML文件
        
        Args:
            file_path: PlantUML文件路径
            
        Returns:
            Optional[str]: 输出文件路径，失败返回None
        """
        try:
            input_file = Path(file_path)
            output_file = Path(self._generate_output_path(file_path, '.png'))
            
            self.logger.info(f"开始转换PlantUML文件: {input_file} -> {output_file}")
            
            # 执行转换
            success = self._execute_plantuml_command(str(input_file), str(output_file))
            
            if success and output_file.exists():
                self.logger.info(f"PlantUML转换成功: {output_file}")
                return str(output_file)
            else:
                self.logger.error(f"PlantUML转换失败: {input_file}, success={success}, file_exists={output_file.exists()}")
                return None
                
        except Exception as e:
            self._handle_conversion_error(e, file_path)
            return None
    
    def _check_dependencies(self) -> DependencyStatus:
        """
        检查PlantUML转换依赖
        
        Returns:
            DependencyStatus: 依赖检查结果
        """
        if self._dependency_status is not None:
            return self._dependency_status
        
        self.logger.info("检查PlantUML转换依赖...")
        
        # 检查Java
        java_available, java_version = self._check_java_availability()
        
        # 检查PlantUML JAR
        plantuml_jar_path = self._get_plantuml_jar_path()
        plantuml_version = None
        if plantuml_jar_path:
            plantuml_version = self._get_plantuml_version(plantuml_jar_path)
        
        # 检查Graphviz
        graphviz_available, graphviz_version = self._check_graphviz_availability()
        
        self._dependency_status = DependencyStatus(
            java_available=java_available,
            java_version=java_version,
            plantuml_jar_path=plantuml_jar_path,
            plantuml_version=plantuml_version,
            graphviz_available=graphviz_available,
            graphviz_version=graphviz_version
        )
        
        # 记录依赖状态
        self._log_dependency_status(self._dependency_status)
        
        return self._dependency_status
    
    def _check_java_availability(self) -> tuple[bool, Optional[str]]:
        """检查Java是否可用"""
        try:
            result = subprocess.run(
                ['java', '-version'],
                capture_output=True,
                text=True,
                timeout=10
            )
            
            if result.returncode == 0:
                # Java版本信息通常在stderr中
                version_output = result.stderr or result.stdout
                version_match = re.search(r'version "([^"]+)"', version_output)
                version = version_match.group(1) if version_match else "未知版本"
                return True, version
            else:
                return False, None
                
        except (subprocess.TimeoutExpired, FileNotFoundError, Exception):
            return False, None
    
    def _get_plantuml_jar_path(self) -> Optional[str]:
        """
        获取PlantUML JAR包路径
        
        查找顺序：
        1. 配置文件指定路径
        2. 环境变量PLANTUML_JAR
        3. 当前目录
        4. 系统PATH中的plantuml.jar
        5. 用户主目录/.plantuml/
        
        Returns:
            Optional[str]: JAR包路径，未找到返回None
        """
        if self._jar_path_cache:
            return self._jar_path_cache
        
        search_paths = []
        
        # 1. 配置文件指定路径
        if self.plantuml_config.plantuml_jar_path:
            search_paths.append(self.plantuml_config.plantuml_jar_path)
        
        # 2. 环境变量
        env_path = os.environ.get('PLANTUML_JAR')
        if env_path:
            search_paths.append(env_path)
        
        # 3. 当前目录
        search_paths.append(os.path.join(os.getcwd(), self.PLANTUML_JAR_NAME))
        
        # 4. 系统PATH中查找
        system_jar = shutil.which('plantuml.jar')
        if system_jar:
            search_paths.append(system_jar)
        
        # 5. 用户主目录
        home_path = os.path.join(os.path.expanduser('~'), '.plantuml', self.PLANTUML_JAR_NAME)
        search_paths.append(home_path)
        
        # 6. 项目目录下的tools文件夹
        project_tools_path = os.path.join(os.path.dirname(__file__), '..', '..', 'tools', self.PLANTUML_JAR_NAME)
        search_paths.append(os.path.abspath(project_tools_path))
        
        # 查找JAR文件
        for path in search_paths:
            if os.path.isfile(path):
                self._jar_path_cache = path
                self.logger.info(f"找到PlantUML JAR: {path}")
                return path
        
        self.logger.warning(f"未找到PlantUML JAR包，搜索路径: {search_paths}")
        return None
    
    def _get_plantuml_version(self, jar_path: str) -> Optional[str]:
        """获取PlantUML版本信息"""
        try:
            result = subprocess.run(
                ['java', '-jar', jar_path, '-version'],
                capture_output=True,
                text=True,
                timeout=10
            )
            
            if result.returncode == 0:
                output = result.stdout or result.stderr
                # 提取版本号
                version_match = re.search(r'PlantUML version ([\d\.]+)', output)
                return version_match.group(1) if version_match else "未知版本"
            
        except Exception:
            pass
        
        return None
    
    def _check_graphviz_availability(self) -> tuple[bool, Optional[str]]:
        """检查Graphviz是否可用"""
        try:
            result = subprocess.run(
                ['dot', '-V'],
                capture_output=True,
                text=True,
                timeout=10
            )
            
            if result.returncode == 0:
                # Graphviz版本信息通常在stderr中
                version_output = result.stderr or result.stdout
                version_match = re.search(r'dot - graphviz version ([\d\.]+)', version_output)
                version = version_match.group(1) if version_match else "未知版本"
                return True, version
            else:
                return False, None
                
        except (subprocess.TimeoutExpired, FileNotFoundError, Exception):
            return False, None
    
    def _execute_plantuml_command(self, input_file: str, output_file: str) -> bool:
        """
        执行PlantUML转换命令
        
        Args:
            input_file: 输入PlantUML文件路径
            output_file: 输出PNG文件路径
            
        Returns:
            bool: 转换是否成功
        """
        try:
            # 构建命令
            command = self._build_plantuml_command(input_file, output_file)
            
            self.logger.debug(f"执行PlantUML命令: {' '.join(command)}")
            
            # 执行命令
            # 当使用shell=True时，需要将命令转换为字符串
            command_str = ' '.join(command)
            result = subprocess.run(
                command_str,
                capture_output=True,
                text=True,
                timeout=self.plantuml_config.timeout,
                shell=True
            )
            
            if result.returncode == 0:
                return True
            else:
                error_msg = self._parse_plantuml_error(result.stderr or result.stdout)
                self.logger.error(f"PlantUML转换失败 (返回码: {result.returncode}): {error_msg}")
                self.logger.debug(f"标准输出: {result.stdout}")
                self.logger.debug(f"标准错误: {result.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            self.logger.error(f"PlantUML转换超时 ({self.plantuml_config.timeout}秒): {input_file}")
            return False
        except Exception as e:
            self.logger.error(f"执行PlantUML命令失败: {e}")
            return False
    
    def _build_plantuml_command(self, input_file: str, output_file: str) -> List[str]:
        """
        构建PlantUML命令行参数
        
        Returns:
            List[str]: 命令行参数列表
        """
        command = ['java']
        
        # 添加Java选项
        if self.plantuml_config.java_options:
            command.extend(self.plantuml_config.java_options)
        else:
            # 默认Java选项
            command.extend(['-Xmx1024m', '-Djava.awt.headless=true'])
        
        # 添加JAR包 (使用引号包围路径以处理空格)
        jar_path = self._dependency_status.plantuml_jar_path
        if ' ' in jar_path:
            command.extend(['-jar', f'"{jar_path}"'])
        else:
            command.extend(['-jar', jar_path])
        
        # 添加PlantUML选项
        command.extend([
            f'-t{self.plantuml_config.format}',  # 输出格式
            f'-charset', self.plantuml_config.charset,  # 字符编码
        ])
        
        # 如果指定了Graphviz路径
        if self.plantuml_config.graphviz_dot_path:
            command.extend(['-graphvizdot', self.plantuml_config.graphviz_dot_path])
        
        # 输出目录 (使用引号包围路径以处理空格)
        output_dir = os.path.dirname(output_file)
        if ' ' in output_dir:
            command.extend(['-o', f'"{output_dir}"'])
        else:
            command.extend(['-o', output_dir])
        
        # 输入文件 (使用引号包围路径以处理空格)
        if ' ' in input_file:
            command.append(f'"{input_file}"')
        else:
            command.append(input_file)
        
        return command
    
    def _parse_plantuml_error(self, error_output: str) -> str:
        """
        解析PlantUML错误输出，提供友好的错误信息
        
        Args:
            error_output: PlantUML命令的错误输出
            
        Returns:
            str: 格式化的错误信息
        """
        if not error_output:
            return "未知错误"
        
        # 常见错误模式
        error_patterns = {
            r'Syntax error': '语法错误',
            r'Cannot find Graphviz': 'Graphviz未安装或未找到',
            r'OutOfMemoryError': '内存不足，请增加Java堆内存',
            r'FileNotFoundException': '文件未找到',
            r'AccessDeniedException': '文件访问权限不足',
        }
        
        for pattern, friendly_msg in error_patterns.items():
            if re.search(pattern, error_output, re.IGNORECASE):
                return f"{friendly_msg}: {error_output.strip()}"
        
        return error_output.strip()
    
    def _handle_conversion_error(self, error: Exception, input_file: str) -> None:
        """
        处理转换错误，记录日志并提供解决建议
        """
        error_msg = str(error)
        self.logger.error(f"转换文件 {input_file} 时发生错误: {error_msg}")
        
        # 提供解决建议
        if "java" in error_msg.lower():
            self.logger.error("建议: 请检查Java是否正确安装并在PATH中")
        elif "plantuml" in error_msg.lower():
            self.logger.error("建议: 请检查PlantUML JAR包是否存在")
        elif "graphviz" in error_msg.lower():
            self.logger.error("建议: 请安装Graphviz并确保dot命令在PATH中")
        elif "timeout" in error_msg.lower():
            self.logger.error(f"建议: 文件可能过大，请增加超时时间（当前: {self.plantuml_config.timeout}秒）")
    
    def _log_dependency_status(self, status: DependencyStatus) -> None:
        """记录依赖状态日志"""
        self.logger.info("=== PlantUML依赖检查结果 ===")
        
        if status.java_available:
            self.logger.info(f"✓ Java: {status.java_version}")
        else:
            self.logger.error("✗ Java: 未安装或不可用")
        
        if status.plantuml_jar_path:
            version_info = f" ({status.plantuml_version})" if status.plantuml_version else ""
            self.logger.info(f"✓ PlantUML JAR: {status.plantuml_jar_path}{version_info}")
        else:
            self.logger.error("✗ PlantUML JAR: 未找到")
        
        if status.graphviz_available:
            self.logger.info(f"✓ Graphviz: {status.graphviz_version}")
        else:
            self.logger.warning("⚠ Graphviz: 未安装（某些图表类型可能无法正常显示）")
        
        if status.is_ready:
            self.logger.info("✓ PlantUML转换环境就绪")
        else:
            self.logger.error("✗ PlantUML转换环境未就绪，请安装缺失的依赖")
        
        self.logger.info("=" * 30)