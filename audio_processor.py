import os
import sys
import json
import re
import subprocess
import shutil
from datetime import datetime
import PySimpleGUI as sg
import time

class AudioProcessor:
    def __init__(self):
        # 配置文件路径 - 标准化确保跨平台兼容性
        self.config_file = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json'))
        
        # 加载配置
        self.load_config()
        
        # 创建GUI界面
        self.create_layout()
        
        # 检查ffmpeg
        self.check_ffmpeg()
        
        # 转换格式窗口
        self.convert_window = None
        
    def load_config(self):
        """加载用户配置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # 标准化上次使用的文件夹路径
                    self.last_folder = os.path.normpath(config.get('last_folder', ''))
                    self.check_missing_files = config.get('check_missing_files', False)
                    # 加载转换配置参数
                    self.convert_config = config.get('convert_config', {
                        'format': 'mp3',
                        'codec': 'libmp3lame',
                        'bitrate': '192k',
                        'channels': '2',
                        'sample_rate': '44100',
                        'start_time': '',
                        'end_time': ''
                    })
            else:
                self.last_folder = ''
                self.check_missing_files = False
                # 默认转换配置
                self.convert_config = {
                    'format': 'mp3',
                    'codec': 'libmp3lame',
                    'bitrate': '192k',
                    'channels': '2',
                    'sample_rate': '44100',
                    'start_time': '',
                    'end_time': ''
                }
        except:
            self.last_folder = ''
            self.check_missing_files = False
            self.convert_config = {
                'format': 'mp3',
                'codec': 'libmp3lame',
                'bitrate': '192k',
                'channels': '2',
                'sample_rate': '44100',
                'start_time': '',
                'end_time': ''
            }
    
    def save_config(self):
        """保存用户配置"""
        config = {
            'last_folder': self.last_folder,
            'check_missing_files': self.check_missing_files,
            'convert_config': self.convert_config
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            sg.popup_error(f"保存配置失败: {str(e)}")
    
    def create_layout(self):
        """创建GUI布局"""
        sg.theme('LightBlue2')  # 设置主题
        
        # 定义常见音频格式
        audio_formats = ['mp3', 'wav', 'flac', 'aac', 'ogg', 'wma', 'm4a', 'ts']
        
        # 创建布局
        layout = [
            [sg.Text('音频处理工具', font=('Arial', 16, 'bold'))],
            [sg.Text('FFmpeg 状态: ', key='-FFMPEG_STATUS-')],
            [sg.HorizontalSeparator()],
            [sg.Text('选择音频文件夹:', size=(15, 1)), 
             sg.InputText(self.last_folder, key='-FOLDER-', size=(40, 1)), 
             sg.FolderBrowse('浏览', key='-BROWSE-')],
            [sg.Checkbox('检查缺失的音频文件（数字序列）', default=self.check_missing_files, key='-CHECK_MISSING-')],
            [sg.Text('音频文件列表:', size=(15, 1))],
            [sg.Multiline(size=(60, 10), key='-FILE_LIST-', disabled=True)],
            [sg.HorizontalSeparator()],
            [sg.Button('扫描文件', key='-SCAN-'), 
             sg.Button('合并音频', key='-MERGE-'), 
             sg.Button('转换格式', key='-CONVERT-')],
            [sg.HorizontalSeparator()],
            [sg.Text('日志:', size=(15, 1))],
            [sg.Multiline(size=(60, 5), key='-LOG-', disabled=True)]
        ]
        
        # 创建窗口
        self.window = sg.Window('音频处理工具', layout, resizable=True, finalize=True)
    
    def check_ffmpeg(self):
        """检查ffmpeg是否安装并显示版本"""
        try:
            # 尝试运行ffmpeg命令
            result = subprocess.run(['ffmpeg', '-version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            
            if result.returncode == 0:
                # 提取ffmpeg版本信息
                version_line = result.stdout.split('\n')[0]
                self.window['-FFMPEG_STATUS-'].update(f'FFmpeg 状态: 已安装 。')
                self.log(f"FFmpeg已安装: {version_line}")
                return True
            else:
                self.window['-FFMPEG_STATUS-'].update('FFmpeg 状态: 未安装')
                self.log("错误: 未找到FFmpeg。请先安装FFmpeg并添加到系统环境变量中。")
                sg.popup_error('未找到FFmpeg。请先安装FFmpeg并添加到系统环境变量中。')
                return False
        except FileNotFoundError:
            self.window['-FFMPEG_STATUS-'].update('FFmpeg 状态: 未安装')
            self.log("错误: 未找到FFmpeg。请先安装FFmpeg并添加到系统环境变量中。")
            sg.popup_error('未找到FFmpeg。请先安装FFmpeg并添加到系统环境变量中。')
            return False
    
    def log(self, message):
        """向日志区域添加消息"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.window['-LOG-'].print(f'[{timestamp}] {message}')
        
    def get_audio_info(self, file_path):
        """使用ffprobe获取音频文件信息"""
        try:
            # 标准化文件路径并转换为FFmpeg兼容格式
            ffmpeg_file_path = file_path.replace('\\', '/')
            
            # 使用ffprobe获取详细信息
            cmd = ['ffprobe', '-v', 'quiet', '-print_format', 'json', 
                  '-show_format', '-show_streams', ffmpeg_file_path]
            result = subprocess.run(cmd, check=False, stdout=subprocess.PIPE, 
                                   stderr=subprocess.PIPE, text=True)
            
            if result.returncode != 0:
                self.log(f"获取音频信息失败: {result.stderr}")
                return None
            
            # 解析JSON输出
            info = json.loads(result.stdout)
            
            # 提取音频流信息
            audio_stream = None
            for stream in info.get('streams', []):
                if stream.get('codec_type') == 'audio':
                    audio_stream = stream
                    break
            
            if not audio_stream:
                self.log(f"未找到音频流: {file_path}")
                return None
            
            # 格式化信息
            format_info = info.get('format', {})
            duration = float(format_info.get('duration', 0))
            bit_rate = format_info.get('bit_rate', '未知')
            
            # 转换时长为分:秒格式
            minutes, seconds = divmod(duration, 60)
            duration_str = f"{int(minutes)}:{int(seconds):02d}"
            
            audio_info = {
                '文件名': os.path.basename(file_path),
                '时长': duration_str,
                '编码方式': audio_stream.get('codec_name', '未知'),
                '码率': f"{int(int(bit_rate)/1000)} kbps" if bit_rate != '未知' else '未知',
                '采样率': f"{audio_stream.get('sample_rate', '未知')} Hz",
                '声道数': audio_stream.get('channels', '未知'),
                '声道布局': audio_stream.get('channel_layout', '未知'),
                '格式': format_info.get('format_name', '未知')
            }
            
            return audio_info
        except Exception as e:
            self.log(f"获取音频信息时发生错误: {str(e)}")
            return None
        
    def convert_format_window(self):
        """创建单独的转换格式页面"""
        sg.theme('LightBlue2')
        
        # 定义音频格式、编码器、声道、采样率选项
        formats = ['mp3', 'wav', 'flac', 'aac', 'ogg', 'wma', 'm4a']
        codecs = {
            'mp3': ['libmp3lame'],
            'wav': ['pcm_s16le', 'pcm_s24le', 'pcm_f32le'],
            'flac': ['flac'],
            'aac': ['aac', 'libfdk_aac'],
            'ogg': ['libvorbis'],
            'wma': ['wmav2'],
            'm4a': ['aac', 'libfdk_aac']
        }
        bitrates = ['96k', '128k', '192k', '256k', '320k']
        channels = ['1', '2', '4', '6']
        sample_rates = ['22050', '44100', '48000', '96000']
        
        # 布局设计
        layout = [
            [sg.Text('音频格式转换', font=('Arial', 16, 'bold'))],
            [sg.HorizontalSeparator()],
            [sg.Text('选择音频文件夹:', size=(15, 1)),
             sg.InputText(self.last_folder, key='-FOLDER-', size=(40, 1)),
             sg.FolderBrowse('浏览', key='-BROWSE-')],
            [sg.Button('扫描文件', key='-SCAN-')],
            [sg.Text('音频文件列表:', size=(15, 1))],
            [sg.Listbox(values=[], size=(60, 8), key='-FILE_LIST-', enable_events=True)],
            [sg.HorizontalSeparator()],
            [sg.Text('音频文件信息:', font=('Arial', 12, 'bold'))],
            [sg.Multiline(size=(60, 8), key='-AUDIO_INFO-', disabled=True)],
            [sg.HorizontalSeparator()],
            [sg.Text('转换设置:', font=('Arial', 12, 'bold'))],
            [sg.Text('输出格式:', size=(15, 1)),
             sg.Combo(formats, default_value=self.convert_config['format'],
                     key='-OUTPUT_FORMAT-', size=(10, 1), enable_events=True)],
            [sg.Text('编码器:', size=(15, 1)),
             sg.Combo(codecs.get(self.convert_config['format'], []),
                     default_value=self.convert_config['codec'],
                     key='-CODEC-', size=(15, 1))],
            [sg.Text('比特率:', size=(15, 1)),
             sg.Combo(bitrates, default_value=self.convert_config['bitrate'],
                     key='-BITRATE-', size=(10, 1))],
            [sg.Text('声道数:', size=(15, 1)),
             sg.Combo(channels, default_value=self.convert_config['channels'],
                     key='-CHANNELS-', size=(10, 1))],
            [sg.Text('采样率:', size=(15, 1)),
             sg.Combo(sample_rates, default_value=self.convert_config['sample_rate'],
                     key='-SAMPLE_RATE-', size=(10, 1))],
            [sg.Text('起始时间:', size=(15, 1)),
             sg.InputText(self.convert_config['start_time'], key='-START_TIME-',
                         size=(15, 1), tooltip='格式: HH:MM:SS')],
            [sg.Text('结束时间:', size=(15, 1)),
             sg.InputText(self.convert_config['end_time'], key='-END_TIME-',
                         size=(15, 1), tooltip='格式: HH:MM:SS')],
            [sg.HorizontalSeparator()],
            [sg.Button('转换选中文件', key='-CONVERT_SELECTED-'),
             sg.Button('转换所有文件', key='-CONVERT_ALL-'),
             sg.Button('保存配置', key='-SAVE_CONFIG-'),
             sg.Button('关闭', key='-CLOSE-')]
        ]
        
        # 创建窗口
        self.convert_window = sg.Window('音频格式转换', layout, resizable=True, finalize=True)
        
        # 保存当前选中的文件夹路径
        current_folder = self.last_folder
        audio_files = []
        
        # 窗口事件循环
        while True:
            event, values = self.convert_window.read()
            
            if event == sg.WIN_CLOSED or event == '-CLOSE-':
                break
            
            # 扫描文件夹
            if event == '-SCAN-':
                folder_path = values['-FOLDER-']
                if not folder_path:
                    sg.popup_error('请先选择文件夹！')
                    continue
                
                # 标准化文件夹路径
                folder_path = os.path.normpath(folder_path)
                current_folder = folder_path
                self.last_folder = folder_path
                
                # 扫描音频文件
                audio_files = self.scan_folder(folder_path)
                self.convert_window['-FILE_LIST-'].update(audio_files)
                self.log(f"转换页面 - 找到 {len(audio_files)} 个音频文件")
            
            # 文件列表选择变化
            if event == '-FILE_LIST-':
                if values['-FILE_LIST-']:
                    selected_file = values['-FILE_LIST-'][0]
                    file_path = os.path.join(current_folder, selected_file)
                    
                    # 获取并显示音频信息
                    audio_info = self.get_audio_info(file_path)
                    if audio_info:
                        info_text = '\n'.join([f"{key}: {value}" for key, value in audio_info.items()])
                        self.convert_window['-AUDIO_INFO-'].update(info_text)
            
            # 输出格式变化，更新编码器选项
            if event == '-OUTPUT_FORMAT-':
                output_format = values['-OUTPUT_FORMAT-']
                self.convert_window['-CODEC-'].update(values=codecs.get(output_format, []))
                # 如果当前选择的编码器不适用于新格式，则选择第一个可用编码器
                current_codec = values['-CODEC-']
                available_codecs = codecs.get(output_format, [])
                if current_codec not in available_codecs and available_codecs:
                    self.convert_window['-CODEC-'].update(set_to_index=0)
            
            # 保存配置
            if event == '-SAVE_CONFIG-':
                # 更新转换配置
                self.convert_config = {
                    'format': values['-OUTPUT_FORMAT-'],
                    'codec': values['-CODEC-'],
                    'bitrate': values['-BITRATE-'],
                    'channels': values['-CHANNELS-'],
                    'sample_rate': values['-SAMPLE_RATE-'],
                    'start_time': values['-START_TIME-'],
                    'end_time': values['-END_TIME-']
                }
                # 保存配置
                self.save_config()
                sg.popup('配置保存成功！')
                self.log("转换配置已保存")
            
            # 转换选中文件
            if event == '-CONVERT_SELECTED-':
                if not audio_files:
                    sg.popup_error('没有找到音频文件！')
                    continue
                
                if not values['-FILE_LIST-']:
                    sg.popup_error('请先选择要转换的文件！')
                    continue
                
                selected_files = values['-FILE_LIST-']
                self.perform_conversion(current_folder, selected_files, values)
            
            # 转换所有文件
            if event == '-CONVERT_ALL-':
                if not audio_files:
                    sg.popup_error('没有找到音频文件！')
                    continue
                
                self.perform_conversion(current_folder, audio_files, values)
        
        # 关闭窗口
        self.convert_window.close()
        self.convert_window = None
    
    def scan_folder(self, folder_path):
        """扫描文件夹中的音频文件"""
        # 标准化文件夹路径
        folder_path = os.path.normpath(folder_path)
        
        if not os.path.exists(folder_path):
            sg.popup_error('文件夹不存在！')
            return []
        
        # 保存最后选择的文件夹
        self.last_folder = folder_path
        
        # 常见音频文件扩展名
        audio_extensions = ['.mp3', '.wav', '.flac', '.aac', '.ogg', '.wma', '.m4a', '.ts']
        
        audio_files = []
        
        # 遍历文件夹
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if os.path.isfile(file_path):
                ext = os.path.splitext(file)[1].lower()
                if ext in audio_extensions:
                    audio_files.append(file)
        
        # 按文件名排序（尝试按数字排序）
        try:
            audio_files.sort(key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else x)
        except:
            audio_files.sort()
        
        return audio_files
    
    def check_missing_audio_files(self, folder_path, audio_files):
        """检查缺失的音频文件（假设文件名是数字序列）"""
        if not audio_files:
            return []
        
        # 标准化文件夹路径
        folder_path = os.path.normpath(folder_path)
        
        # 提取文件名中的数字
        numbers = []
        file_patterns = []
        
        for file in audio_files:
            # 找出文件名中的所有数字序列
            all_matches = re.findall(r'\d+', file)
            if all_matches:
                # 优先使用最长的数字序列，通常这是主要的序号
                longest_match = max(all_matches, key=len)
                try:
                    num = int(longest_match)
                    numbers.append(num)
                    # 保存文件名模式以便后续生成缺失文件名
                    file_patterns.append((file, longest_match))
                except:
                    pass
        
        if not numbers:
            return []
        
        # 找到最小和最大数字
        min_num = min(numbers)
        max_num = max(numbers)
        
        # 找出缺失的数字
        missing_numbers = []
        for i in range(min_num, max_num + 1):
            if i not in numbers:
                missing_numbers.append(i)
        
        # 如果有缺失的数字，生成txt文件
        if missing_numbers:
            missing_file_path = os.path.join(folder_path, 'missing_files.txt')
            # 标准化输出文件路径
            missing_file_path = os.path.normpath(missing_file_path)
            with open(missing_file_path, 'w', encoding='utf-8') as f:
                for num in missing_numbers:
                    # 尝试生成与现有文件相同格式的文件名
                    # 使用保存的模式信息
                    for file, original_num in file_patterns:
                        # 替换原始数字为缺失的数字
                        missing_file_name = file.replace(original_num, str(num))
                        f.write(f"{missing_file_name}\n")
                        break
            self.log(f"已生成缺失文件清单: {missing_file_path}")
        
        return missing_numbers
    
    def create_ffmpeg_file_list(self, folder_path, audio_files):
        """创建ffmpeg合并文件列表"""
        file_list_path = os.path.join(folder_path, 'files.txt')
        
        # 确保路径在FFmpeg中兼容（使用正斜杠）
        try:
            with open(file_list_path, 'w', encoding='utf-8') as f:
                self.log(f"创建文件列表: {file_list_path}")
                for file in audio_files:
                    # 将路径转换为FFmpeg兼容的格式
                    full_path = os.path.join(folder_path, file)
                    # 标准化路径，确保跨平台兼容性
                    full_path = os.path.normpath(full_path)
                    # 替换所有反斜杠为正斜杠，这在所有平台上对FFmpeg都有效
                    full_path = full_path.replace('\\', '/')
                    
                    # 写入符合FFmpeg concat协议格式的路径
                    # 使用双引号替代单引号，防止路径中包含单引号时出现问题
                    f.write(f"file '{full_path}'\n")
                    self.log(f"添加文件到列表: {full_path}")
            
            # 验证文件列表是否成功创建
            if os.path.exists(file_list_path):
                file_list_size = os.path.getsize(file_list_path)
                self.log(f"文件列表创建成功，大小: {file_list_size} 字节")
                # 读取前几行日志，以便调试
                with open(file_list_path, 'r', encoding='utf-8') as f:
                    first_lines = f.readlines()[:5]
                    self.log(f"文件列表前几行: {''.join(first_lines).strip()}")
            else:
                self.log("错误: 文件列表未成功创建")
                raise Exception("创建文件列表失败")
                
        except Exception as e:
            self.log(f"创建文件列表时发生错误: {str(e)}")
            raise
        
        return file_list_path
    
    def merge_audio_files(self, folder_path, audio_files):
        """使用ffmpeg合并音频文件"""
        if not audio_files:
            sg.popup_error('没有找到音频文件！')
            return False
        
        # 标准化文件夹路径
        folder_path = os.path.normpath(folder_path)
        
        # 创建ffmpeg文件列表
        file_list_path = self.create_ffmpeg_file_list(folder_path, audio_files)
        self.log(f"已创建文件列表: {file_list_path}")
        
        # 检测输入音频文件的格式
        if audio_files:
            # 获取第一个音频文件的扩展名
            first_file_ext = os.path.splitext(audio_files[0])[1].lower()
            # 移除点号
            output_format = first_file_ext[1:] if first_file_ext.startswith('.') else first_file_ext
            # 如果没有识别到格式或格式为空，默认使用mp3
            if not output_format:
                output_format = 'mp3'
        else:
            output_format = 'mp3'
        
        # 生成输出文件名，使用与输入文件相同的格式
        output_file = os.path.join(folder_path, f"merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{output_format}")
        # 标准化输出文件路径
        output_file = os.path.normpath(output_file)
        
        self.log(f"输出文件格式: {output_format}")
        
        try:
            # 执行ffmpeg合并命令前，确保所有路径使用正斜杠格式
            # 对于FFmpeg命令参数也需要转换路径格式
            ffmpeg_file_list_path = file_list_path.replace('\\', '/')
            ffmpeg_output_file = output_file.replace('\\', '/')
            
            self.log("开始合并音频文件...")
            # 使用无损合并方式，保持输入和输出音频格式一致
            cmd = ['ffmpeg', '-f', 'concat', '-safe', '0', '-i', ffmpeg_file_list_path, '-c', 'copy', ffmpeg_output_file]
            self.log(f"执行FFmpeg无损合并命令: {' '.join(cmd)}")
            result = subprocess.run(cmd, check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            
            if result.returncode != 0:
                self.log(f"FFmpeg无损合并失败输出: {result.stderr}")
                self.log(f"FFmpeg标准输出: {result.stdout}")
                
                # 询问用户是否尝试使用重新编码方式合并
                if sg.popup_yes_no('无损合并失败！这可能是由于音频文件编码格式不一致导致的。\n是否尝试使用重新编码方式进行合并？') == 'Yes':
                    self.log("用户选择尝试重新编码合并方式")
                    # 使用重新编码方式合并
                    # 根据输出格式选择合适的编码器
                    if output_format == 'mp3':
                        encoder_cmd = ['-c:a', 'libmp3lame', '-q:a', '2']
                    elif output_format in ['wav', 'flac']:
                        # 对于无损格式，使用适当的编码器和参数
                        encoder_cmd = ['-c:a', 'pcm_s16le'] if output_format == 'wav' else ['-c:a', 'flac']
                    else:
                        # 默认使用通用编码设置
                        encoder_cmd = ['-c:a', 'aac', '-b:a', '192k']
                    
                    cmd = ['ffmpeg', '-f', 'concat', '-safe', '0', '-i', ffmpeg_file_list_path] + encoder_cmd + [ffmpeg_output_file]
                    self.log(f"执行FFmpeg重新编码合并命令: {' '.join(cmd)}")
                    result = subprocess.run(cmd, check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                    
                    if result.returncode != 0:
                        self.log(f"FFmpeg重新编码合并失败输出: {result.stderr}")
                        raise Exception(f"FFmpeg重新编码合并失败，退出代码: {result.returncode}\n错误输出: {result.stderr}")
                else:
                    raise Exception(f"FFmpeg无损合并失败，退出代码: {result.returncode}\n错误输出: {result.stderr}")
            
            self.log(f"音频文件合并成功: {output_file}")
            
            # 删除合并前的音频文件
            if sg.popup_yes_no('音频文件合并成功！是否删除原始音频文件？') == 'Yes':
                for file in audio_files:
                    try:
                        os.remove(os.path.join(folder_path, file))
                        self.log(f"已删除: {file}")
                    except Exception as e:
                        self.log(f"删除文件失败 {file}: {str(e)}")
            
            return True
        except Exception as e:
            self.log(f"合并音频文件失败: {str(e)}")
            sg.popup_error(f"合并音频文件失败: {str(e)}")
            return False
    
    def convert_audio_format(self, folder_path, audio_files, output_format):
        """转换音频文件格式（兼容旧接口）"""
        # 创建一个临时的values字典来传递参数
        values = {
            '-OUTPUT_FORMAT-': output_format,
            '-CODEC-': 'libmp3lame' if output_format == 'mp3' else 'aac',
            '-BITRATE-': '192k',
            '-CHANNELS-': '2',
            '-SAMPLE_RATE-': '44100',
            '-START_TIME-': '',
            '-END_TIME-': ''
        }
        return self.perform_conversion(folder_path, audio_files, values)
    
    def perform_conversion(self, folder_path, audio_files, values):
        """执行音频转换，支持自定义参数"""
        if not audio_files:
            sg.popup_error('没有找到音频文件！')
            return False
        
        # 标准化文件夹路径
        folder_path = os.path.normpath(folder_path)
        
        # 获取转换参数
        output_format = values['-OUTPUT_FORMAT-']
        codec = values['-CODEC-']
        bitrate = values['-BITRATE-']
        channels = values['-CHANNELS-']
        sample_rate = values['-SAMPLE_RATE-']
        start_time = values['-START_TIME-']
        end_time = values['-END_TIME-']
        
        # 创建输出文件夹
        output_folder = os.path.join(folder_path, f"converted_{output_format}")
        os.makedirs(output_folder, exist_ok=True)
        

        
        success_count = 0
        
        try:
            for file in audio_files:
                input_file = os.path.join(folder_path, file)
                # 标准化输入文件路径
                input_file = os.path.normpath(input_file)
                
                base_name = os.path.splitext(file)[0]
                output_file = os.path.join(output_folder, f"{base_name}.{output_format}")
                # 标准化输出文件路径
                output_file = os.path.normpath(output_file)
                
                self.log(f"正在转换: {file} -> {output_file}")
                self.log(f"转换参数 - 编码器: {codec}, 比特率: {bitrate}, 声道: {channels}, 采样率: {sample_rate}")
                
                # 执行ffmpeg转换命令前，确保所有路径使用正斜杠格式
                ffmpeg_input_file = input_file.replace('\\', '/')
                ffmpeg_output_file = output_file.replace('\\', '/')
                
                # 构建命令参数
                cmd = ['ffmpeg', '-i', ffmpeg_input_file]
                
                # 添加时间裁剪参数
                if start_time:
                    cmd.extend(['-ss', start_time])
                    self.log(f"应用起始时间: {start_time}")
                if end_time:
                    cmd.extend(['-to', end_time])
                    self.log(f"应用结束时间: {end_time}")
                
                # 添加音频编码参数
                cmd.extend(['-c:a', codec])
                cmd.extend(['-b:a', bitrate])
                cmd.extend(['-ac', channels])
                cmd.extend(['-ar', sample_rate])
                
                # 添加覆盖输出参数和输出文件路径
                cmd.extend(['-y', ffmpeg_output_file])
                
                # 执行转换命令
                self.log(f"执行FFmpeg转换命令: {' '.join(cmd)}")
                result = subprocess.run(cmd, check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                
                if result.returncode != 0:
                    self.log(f"FFmpeg错误输出: {result.stderr}")
                    self.log(f"FFmpeg标准输出: {result.stdout}")
                    raise Exception(f"FFmpeg命令失败，退出代码: {result.returncode}\n错误输出: {result.stderr}")
                
                success_count += 1
                self.log(f"转换成功: {file}")
            
            self.log(f"格式转换完成，成功转换 {success_count}/{len(audio_files)} 个文件")
            sg.popup(f"格式转换完成，成功转换 {success_count}/{len(audio_files)} 个文件\n输出文件夹: {output_folder}")
            
            # 如果转换窗口存在，更新配置
            if self.convert_window:
                self.convert_config = {
                    'format': output_format,
                    'codec': codec,
                    'bitrate': bitrate,
                    'channels': channels,
                    'sample_rate': sample_rate,
                    'start_time': start_time,
                    'end_time': end_time
                }
                self.save_config()
            
            return True
        except Exception as e:
            self.log(f"格式转换失败: {str(e)}")
            sg.popup_error(f"格式转换失败: {str(e)}")
            return False
    
    def run(self):
        """运行应用程序"""
        while True:
            event, values = self.window.read()
            
            if event == sg.WIN_CLOSED:
                # 保存配置并退出
                self.save_config()
                break
            
            if event == '-SCAN-':
                folder_path = values['-FOLDER-']
                if not folder_path:
                    sg.popup_error('请先选择文件夹！')
                    continue
                
                # 保存检查缺失文件选项
                self.check_missing_files = values['-CHECK_MISSING-']
                
                # 扫描文件夹
                self.log(f"开始扫描文件夹: {folder_path}")
                audio_files = self.scan_folder(folder_path)
                
                # 显示文件列表
                self.window['-FILE_LIST-'].update('\n'.join(audio_files))
                self.log(f"找到 {len(audio_files)} 个音频文件")
                
                # 检查缺失文件
                if self.check_missing_files and audio_files:
                    missing_numbers = self.check_missing_audio_files(folder_path, audio_files)
                    if missing_numbers:
                        self.log(f"发现 {len(missing_numbers)} 个缺失的音频文件")
                    else:
                        self.log("未发现缺失的音频文件")
            
            if event == '-MERGE-':
                folder_path = values['-FOLDER-']
                if not folder_path:
                    sg.popup_error('请先选择文件夹！')
                    continue
                
                # 重新扫描文件夹确保文件列表最新
                audio_files = self.scan_folder(folder_path)
                if audio_files:
                    self.merge_audio_files(folder_path, audio_files)
            
            if event == '-CONVERT-':
                # 打开单独的转换格式页面
                self.convert_format_window()
        
        self.window.close()

# 主程序入口
if __name__ == '__main__':
    try:
        app = AudioProcessor()
        app.run()
    except Exception as e:
        sg.popup_error(f"程序运行出错: {str(e)}")
        sys.exit(1)