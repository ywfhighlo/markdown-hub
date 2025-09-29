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
class BatikConfig:
    """Batik转换配置"""
    dpi: int = 300
    width: Optional[int] = None
    height: Optional[int] = None
    quality: float = 1.0
    batik_jar_path: Optional[str] = None
    java_options: List[str] = field(default_factory=list)
    timeout: int = 60  # 转换超时时间（秒）


@dataclass
class BatikDependencyStatus:
    """Batik依赖检查状态"""
    java_available: bool
    java_version: Optional[str]
    batik_jar_path: Optional[str]
    batik_lib_path: Optional[str]
    
    @property
    def is_ready(self) -> bool:
        """是否准备就绪"""
        return self.java_available and bool(self.batik_jar_path) and bool(self.batik_lib_path)


class BatikConverter(BaseConverter):
    """
    Batik SVG到PNG转换器
    继承BaseConverter，专门处理SVG文件转换
    """
    
    # 支持的文件扩展名
    SUPPORTED_EXTENSIONS = ['.svg']
    
    # Batik JAR文件名
    BATIK_ALL_JAR_NAME = 'batik-all.jar'
    BATIK_LIB_DIR_NAME = 'batik-lib'
    
    # 默认配置
    DEFAULT_CONFIG = {
        'dpi': 300,
        'quality': 1.0,
        'timeout': 60
    }

    def __init__(self, output_dir: str, **kwargs):
        super().__init__(output_dir, **kwargs)
        
        # 过滤BatikConfig支持的参数
        batik_config_keys = {
            'dpi', 'width', 'height', 'quality', 'batik_jar_path', 
            'java_options', 'timeout'
        }
        
        # 合并默认配置，只包含BatikConfig支持的参数
        config = {**self.DEFAULT_CONFIG}
        for key, value in kwargs.items():
            if key in batik_config_keys:
                config[key] = value
        
        self.batik_config = BatikConfig(**config)
        
        # 依赖状态缓存
        self._dependency_status: Optional[BatikDependencyStatus] = None
        self._jar_path_cache: Optional[str] = None
        self._lib_path_cache: Optional[str] = None
        
        # 检查依赖
        self._check_dependencies()

    def convert(self, input_path: str) -> List[str]:
        """
        转换SVG文件为PNG
        
        Args:
            input_path: 输入的SVG文件或包含SVG文件的目录
            
        Returns:
            List[str]: 生成的输出文件路径列表
        """
        # 验证输入
        if not self._is_valid_input(input_path, self.SUPPORTED_EXTENSIONS):
            raise ValueError(f"无效的输入文件或目录: {input_path}")
        
        # 检查依赖是否就绪
        if not self._dependency_status or not self._dependency_status.is_ready:
            raise RuntimeError("Batik转换依赖未就绪，请检查Java环境和Batik JAR包")
        
        output_files = []
        
        if os.path.isfile(input_path):
            # 单文件转换
            result = self._convert_single_file(input_path)
            if result:
                output_files.append(result)
        else:
            # 目录批量转换
            svg_files = self._get_files_by_extension(input_path, self.SUPPORTED_EXTENSIONS)
            for file_path in svg_files:
                try:
                    result = self._convert_single_file(file_path)
                    if result:
                        output_files.append(result)
                except Exception as e:
                    self.logger.error(f"转换文件 {file_path} 失败: {e}")
                    continue
        
        self.logger.info(f"Batik转换完成，共生成 {len(output_files)} 个文件")
        return output_files

    def convert_to_file(self, input_path: str, output_path: str, **kwargs) -> tuple[bool, str]:
        """
        转换SVG文件到指定的输出路径
        
        Args:
            input_path: 输入SVG文件路径
            output_path: 输出PNG文件路径
            **kwargs: 额外参数（如dpi等）
            
        Returns:
            tuple[bool, str]: (是否成功, 消息)
        """
        try:
            # 检查依赖是否就绪
            if not self._dependency_status or not self._dependency_status.is_ready:
                return False, "Batik转换依赖未就绪，请检查Java环境和Batik JAR包"
            
            # 验证输入文件
            if not os.path.isfile(input_path):
                return False, f"输入文件不存在: {input_path}"
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            
            # 预处理SVG文件，修复Batik不兼容的语法
            processed_svg_path = self._preprocess_svg_for_batik(input_path)
            
            try:
                # 执行转换
                success = self._execute_batik_command(processed_svg_path, output_path)
                
                if success:
                    return True, f"转换成功: {input_path} -> {output_path}"
                else:
                    return False, f"转换失败: {input_path}"
            finally:
                # 清理临时文件
                if processed_svg_path != input_path and os.path.exists(processed_svg_path):
                    try:
                        os.unlink(processed_svg_path)
                    except:
                        pass
                
        except Exception as e:
            error_msg = f"转换过程中发生错误: {str(e)}"
            self.logger.error(error_msg)
            return False, error_msg

    def _convert_single_file(self, file_path: str) -> Optional[str]:
        """
        转换单个SVG文件
        
        Args:
            file_path: SVG文件路径
            
        Returns:
            Optional[str]: 输出文件路径，失败返回None
        """
        try:
            input_file = Path(file_path)
            output_file = Path(self._generate_output_path(file_path, '.png'))
            
            self.logger.info(f"开始转换SVG文件: {input_file} -> {output_file}")
            
            # 预处理SVG文件
            processed_svg_path = self._preprocess_svg_for_batik(str(input_file))
            
            try:
                # 执行转换
                success = self._execute_batik_command(processed_svg_path, str(output_file))
                
                if success and output_file.exists():
                    self.logger.info(f"Batik转换成功: {output_file}")
                    return str(output_file)
                else:
                    self.logger.error(f"Batik转换失败: {input_file}")
                    return None
            finally:
                # 清理临时文件
                if processed_svg_path != str(input_file) and os.path.exists(processed_svg_path):
                    try:
                        os.unlink(processed_svg_path)
                    except:
                        pass
                
        except Exception as e:
            self._handle_conversion_error(e, file_path)
            return None

    def _check_dependencies(self) -> BatikDependencyStatus:
        """
        检查Batik转换依赖
        
        Returns:
            BatikDependencyStatus: 依赖检查结果
        """
        if self._dependency_status is not None:
            return self._dependency_status
        
        self.logger.info("检查Batik转换依赖...")
        
        # 检查Java
        java_available, java_version = self._check_java_availability()
        
        # 检查Batik JAR和lib目录
        batik_jar_path = self._get_batik_jar_path()
        batik_lib_path = self._get_batik_lib_path()
        
        self._dependency_status = BatikDependencyStatus(
            java_available=java_available,
            java_version=java_version,
            batik_jar_path=batik_jar_path,
            batik_lib_path=batik_lib_path
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

    def _get_batik_jar_path(self) -> Optional[str]:
        """
        获取Batik JAR包路径
        
        查找顺序：
        1. 配置文件指定路径
        2. 环境变量BATIK_JAR
        3. 项目目录下的tools/batik-lib/batik-all.jar
        4. 当前目录
        
        Returns:
            Optional[str]: JAR包路径，未找到返回None
        """
        if self._jar_path_cache:
            return self._jar_path_cache
        
        search_paths = []
        
        # 1. 配置文件指定路径
        if self.batik_config.batik_jar_path:
            search_paths.append(self.batik_config.batik_jar_path)
        
        # 2. 环境变量
        env_path = os.environ.get('BATIK_JAR')
        if env_path:
            search_paths.append(env_path)
        
        # 3. 项目目录下的tools/batik-lib/batik-all.jar
        project_tools_path = os.path.join(os.path.dirname(__file__), '..', '..', 'tools', self.BATIK_LIB_DIR_NAME, self.BATIK_ALL_JAR_NAME)
        search_paths.append(os.path.abspath(project_tools_path))
        
        # 4. 当前目录
        search_paths.append(os.path.join(os.getcwd(), self.BATIK_ALL_JAR_NAME))
        
        # 查找JAR文件
        for path in search_paths:
            if os.path.isfile(path):
                self._jar_path_cache = path
                self.logger.info(f"找到Batik JAR: {path}")
                return path
        
        self.logger.warning(f"未找到Batik JAR包，搜索路径: {search_paths}")
        return None

    def _get_batik_lib_path(self) -> Optional[str]:
        """
        获取Batik lib目录路径
        
        Returns:
            Optional[str]: lib目录路径，未找到返回None
        """
        if self._lib_path_cache:
            return self._lib_path_cache
        
        # 基于JAR包路径推断lib目录
        jar_path = self._get_batik_jar_path()
        if jar_path:
            lib_dir = os.path.dirname(jar_path)
            if os.path.isdir(lib_dir):
                # 检查必要的依赖jar是否存在
                required_jars = [
                    'batik-all.jar',
                    'xmlgraphics-commons-2.11.jar',
                    'xml-apis-1.4.01.jar',
                    'xml-apis-ext-1.3.04.jar'
                ]
                
                all_exist = True
                for jar_name in required_jars:
                    jar_path_check = os.path.join(lib_dir, jar_name)
                    if not os.path.isfile(jar_path_check):
                        all_exist = False
                        break
                
                if all_exist:
                    self._lib_path_cache = lib_dir
                    self.logger.info(f"找到Batik lib目录: {lib_dir}")
                    return lib_dir
        
        self.logger.warning("未找到完整的Batik lib目录")
        return None

    def _execute_batik_command(self, input_file: str, output_file: str) -> bool:
        """
        执行Batik转换命令
        
        Args:
            input_file: 输入SVG文件路径
            output_file: 输出PNG文件路径
            
        Returns:
            bool: 转换是否成功
        """
        try:
            # 构建命令
            command = self._build_batik_command(input_file, output_file)
            
            self.logger.debug(f"执行Batik命令: {' '.join(command)}")
            
            # 执行命令
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=self.batik_config.timeout,
                cwd=os.path.dirname(os.path.abspath(input_file)),
                shell=True  # 在Windows上需要shell=True
            )
            
            # 详细记录输出信息
            self.logger.info(f"Batik返回码: {result.returncode}")
            if result.stdout:
                self.logger.info(f"Batik标准输出: {result.stdout}")
            if result.stderr:
                self.logger.info(f"Batik错误输出: {result.stderr}")
            
            if result.returncode == 0:
                # 检查输出文件是否实际生成
                # Batik会在输出目录中生成与输入文件同名的PNG文件
                input_filename = os.path.splitext(os.path.basename(input_file))[0] + '.png'
                output_dir = os.path.dirname(output_file)
                actual_output_file = os.path.join(output_dir, input_filename)
                
                # 如果输出目录不存在，创建它
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir, exist_ok=True)
                
                # 检查输出文件是否生成
                # Batik会在输出目录中生成与输入文件同名的PNG文件
                input_basename = os.path.splitext(os.path.basename(input_file))[0]
                expected_output_name = input_basename + '.png'
                expected_output_path = os.path.join(output_dir, expected_output_name)
                
                self.logger.info(f"检查输出文件: {expected_output_path}")
                
                if os.path.exists(expected_output_path):
                    file_size = os.path.getsize(expected_output_path)
                    self.logger.info(f"找到输出文件: {expected_output_path}, 大小: {file_size} 字节")
                    if file_size > 0:
                        # 如果输出文件名与期望的不同，重命名它
                        if expected_output_path != output_file:
                            try:
                                # 确保目标目录存在
                                target_dir = os.path.dirname(output_file)
                                if target_dir and not os.path.exists(target_dir):
                                    os.makedirs(target_dir, exist_ok=True)
                                
                                # 移动文件到目标位置
                                import shutil
                                shutil.move(expected_output_path, output_file)
                                self.logger.info(f"重命名输出文件: {expected_output_path} -> {output_file}")
                            except Exception as e:
                                self.logger.error(f"重命名文件失败: {e}")
                                return False
                        return True
                    else:
                        self.logger.error("输出文件为空")
                        return False
                elif os.path.exists(output_file):
                    file_size = os.path.getsize(output_file)
                    self.logger.info(f"输出文件: {output_file}, 大小: {file_size} 字节")
                    return file_size > 0
                else:
                    self.logger.error(f"输出文件未生成，期望: {expected_output_path} 或 {output_file}")
                    # 列出输出目录的内容以便调试
                    if os.path.exists(output_dir):
                        files = os.listdir(output_dir)
                        self.logger.error(f"输出目录内容: {files}")
                        # 检查是否有任何PNG文件
                        png_files = [f for f in files if f.endswith('.png')]
                        if png_files:
                            self.logger.info(f"发现PNG文件: {png_files}")
                            # 尝试使用第一个PNG文件
                            first_png = os.path.join(output_dir, png_files[0])
                            if os.path.getsize(first_png) > 0:
                                try:
                                    import shutil
                                    shutil.move(first_png, output_file)
                                    self.logger.info(f"使用找到的PNG文件: {first_png} -> {output_file}")
                                    return True
                                except Exception as e:
                                    self.logger.error(f"移动PNG文件失败: {e}")
                    return False
            else:
                error_msg = self._parse_batik_error(result.stderr or result.stdout)
                self.logger.error(f"Batik转换失败: {error_msg}")
                return False
                
        except subprocess.TimeoutExpired:
            self.logger.error(f"Batik转换超时 ({self.batik_config.timeout}秒): {input_file}")
            return False
        except Exception as e:
            self.logger.error(f"执行Batik命令失败: {e}")
            return False

    def _build_batik_command(self, input_file: str, output_file: str) -> List[str]:
        """
        构建Batik命令行参数
        
        Returns:
            List[str]: 命令行参数列表
        """
        command = ['java']
        
        # 添加Java选项
        if self.batik_config.java_options:
            command.extend(self.batik_config.java_options)
        else:
            # 默认Java选项
            command.extend(['-Xmx1024m', '-Djava.awt.headless=true'])
        
        # 构建classpath，包含所有必要的jar包
        lib_path = self._dependency_status.batik_lib_path
        classpath_parts = []
        
        # 添加所有jar包到classpath
        required_jars = [
            'batik-all.jar',
            'xmlgraphics-commons-2.11.jar',
            'xml-apis-1.4.01.jar',
            'xml-apis-ext-1.3.04.jar'
        ]
        
        for jar_name in required_jars:
            jar_path = os.path.join(lib_path, jar_name)
            if os.path.isfile(jar_path):
                if ' ' in jar_path:
                    classpath_parts.append(f'"{jar_path}"')
                else:
                    classpath_parts.append(jar_path)
        
        # 设置classpath
        classpath = os.pathsep.join(classpath_parts)
        command.extend(['-cp', classpath])
        
        # 使用自定义的SVG转换器主类
        command.append('org.apache.batik.apps.rasterizer.Main')
        
        # 添加Batik选项
        command.extend(['-m', 'image/png'])  # 输出格式为PNG
        
        # 设置DPI
        if self.batik_config.dpi:
            command.extend(['-dpi', str(self.batik_config.dpi)])
        
        # 设置尺寸
        if self.batik_config.width:
            command.extend(['-w', str(self.batik_config.width)])
        if self.batik_config.height:
            command.extend(['-h', str(self.batik_config.height)])
        
        # 设置质量
        if self.batik_config.quality != 1.0:
            command.extend(['-q', str(self.batik_config.quality)])
        
        # 输出文件
        if ' ' in output_file:
            command.extend(['-d', f'"{output_file}"'])
        else:
            command.extend(['-d', output_file])
        
        # 输入文件
        if ' ' in input_file:
            command.append(f'"{input_file}"')
        else:
            command.append(input_file)
        
        return command

    def _parse_batik_error(self, error_output: str) -> str:
        """
        解析Batik错误输出，提供友好的错误信息
        
        Args:
            error_output: Batik命令的错误输出
            
        Returns:
            str: 格式化的错误信息
        """
        if not error_output:
            return "未知错误"
        
        # 常见错误模式
        error_patterns = {
            r'ClassNotFoundException': '找不到必要的类，请检查jar包依赖',
            r'OutOfMemoryError': '内存不足，请增加Java堆内存',
            r'FileNotFoundException': '文件未找到',
            r'AccessDeniedException': '文件访问权限不足',
            r'SVGException': 'SVG文件格式错误或不支持',
            r'TranscoderException': 'SVG转换过程中发生错误',
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
        elif "batik" in error_msg.lower():
            self.logger.error("建议: 请检查Batik JAR包是否存在")
        elif "timeout" in error_msg.lower():
            self.logger.error(f"建议: 文件可能过大，请增加超时时间（当前: {self.batik_config.timeout}秒）")

    def _preprocess_svg_for_batik(self, input_path: str) -> str:
        """
        预处理SVG文件，修复Batik不兼容的SVG 2.0语法
        
        Args:
            input_path: 原始SVG文件路径
            
        Returns:
            str: 处理后的SVG文件路径（可能是临时文件）
        """
        try:
            # 读取SVG内容
            with open(input_path, 'r', encoding='utf-8') as f:
                svg_content = f.read()
            
            # 检查是否需要修复
            needs_fix = False
            original_content = svg_content
            
            # 修复 orient="auto-start-reverse" -> orient="auto"
            if 'orient="auto-start-reverse"' in svg_content:
                svg_content = svg_content.replace('orient="auto-start-reverse"', 'orient="auto"')
                needs_fix = True
                self.logger.info("修复了 orient='auto-start-reverse' 语法")
            
            # 修复其他可能的SVG 2.0不兼容语法
            # 可以在这里添加更多的修复规则
            
            # 如果不需要修复，直接返回原文件路径
            if not needs_fix:
                return input_path
            
            # 创建临时文件保存修复后的SVG
            with tempfile.NamedTemporaryFile(mode='w', suffix='.svg', delete=False, encoding='utf-8') as f:
                f.write(svg_content)
                temp_path = f.name
            
            self.logger.info(f"创建了预处理的临时SVG文件: {temp_path}")
            return temp_path
            
        except Exception as e:
            self.logger.warning(f"SVG预处理失败，使用原文件: {e}")
            return input_path

    def _log_dependency_status(self, status: BatikDependencyStatus) -> None:
        """记录依赖状态日志"""
        self.logger.info("=== Batik依赖检查结果 ===")
        
        if status.java_available:
            self.logger.info(f"✓ Java: {status.java_version}")
        else:
            self.logger.error("✗ Java: 未安装或不可用")
        
        if status.batik_jar_path:
            self.logger.info(f"✓ Batik JAR: {status.batik_jar_path}")
        else:
            self.logger.error("✗ Batik JAR: 未找到")
        
        if status.batik_lib_path:
            self.logger.info(f"✓ Batik Lib: {status.batik_lib_path}")
        else:
            self.logger.error("✗ Batik Lib: 未找到完整的依赖库")
        
        if status.is_ready:
            self.logger.info("✓ Batik转换环境就绪")
        else:
            self.logger.error("✗ Batik转换环境未就绪，请安装缺失的依赖")
        
        self.logger.info("=" * 30)