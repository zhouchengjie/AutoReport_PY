import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import zipfile
from datetime import datetime, timedelta
import os
from io import BytesIO
import threading
from decimal import Decimal, ROUND_HALF_UP
import subprocess  # 新增：用于调用外部脚本
import sys  # 新增：用于获取Python解释器路径
# 在main.py的顶部添加导入
import sub  # 导入改造后的sub模块

class DataProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("日报数据处理工具")
        self.root.geometry("600x480")  # 调整窗口高度以容纳新按钮
        self.root.resizable(False, False)

        self.input_file = ""
        self.output_file = ""
        self.report_date = datetime.now().strftime("%Y-%m-%d")

        # 变量存储路径
        self.input_zip_path = tk.StringVar()
        self.output_path = tk.StringVar()

        # 创建界面
        self.create_widgets()

    def create_widgets(self):
        """创建GUI界面组件"""
        # 标题
        title_frame = tk.Frame(self.root, bg='#2c3e50', height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)

        title_label = tk.Label(title_frame, text="日报数据处理工具",
                               font=('微软雅黑', 16, 'bold'),
                               bg='#2c3e50', fg='white')
        title_label.pack(expand=True)

        # 主内容区域
        main_frame = tk.Frame(self.root, bg='#ecf0f1')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 输入文件选择
        input_frame = tk.LabelFrame(main_frame, text="输入文件",
                                    font=('微软雅黑', 11, 'bold'),
                                    bg='#ecf0f1', fg='#2c3e50')
        input_frame.pack(fill=tk.X, pady=10)

        tk.Label(input_frame, text="ZIP压缩包路径:",
                 font=('微软雅黑', 10), bg='#ecf0f1').pack(anchor=tk.W, padx=10, pady=5)

        input_path_frame = tk.Frame(input_frame, bg='#ecf0f1')
        input_path_frame.pack(fill=tk.X, padx=10, pady=5)

        self.input_entry = tk.Entry(input_path_frame, textvariable=self.input_zip_path,
                                    font=('微软雅黑', 10), width=50)
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Button(input_path_frame, text="浏览", command=self.select_input_file,
                  font=('微软雅黑', 10), bg='#3498db', fg='white',
                  padx=15).pack(side=tk.RIGHT, padx=(5, 0))

        # 输出路径选择
        output_frame = tk.LabelFrame(main_frame, text="输出路径",
                                     font=('微软雅黑', 11, 'bold'),
                                     bg='#ecf0f1', fg='#2c3e50')
        output_frame.pack(fill=tk.X, pady=10)

        tk.Label(output_frame, text="保存位置:",
                 font=('微软雅黑', 10), bg='#ecf0f1').pack(anchor=tk.W, padx=10, pady=5)

        output_path_frame = tk.Frame(output_frame, bg='#ecf0f1')
        output_path_frame.pack(fill=tk.X, padx=10, pady=5)

        self.output_entry = tk.Entry(output_path_frame, textvariable=self.output_path,
                                     font=('微软雅黑', 10), width=50)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Button(output_path_frame, text="浏览", command=self.select_output_file,
                  font=('微软雅黑', 10), bg='#3498db', fg='white',
                  padx=15).pack(side=tk.RIGHT, padx=(5, 0))

        # 按钮框架 - 新增：调整为可以放两个按钮
        button_frame = tk.Frame(main_frame, bg='#ecf0f1')
        button_frame.pack(fill=tk.X, pady=20)

        # 按钮容器，用于水平排列按钮
        btn_container = tk.Frame(button_frame, bg='#ecf0f1')
        btn_container.pack()

        # 处理按钮
        self.process_button = tk.Button(btn_container, text="处理数据",
                                        command=self.start_processing,
                                        font=('微软雅黑', 12, 'bold'),
                                        bg='#27ae60', fg='white',
                                        width=15, height=2)
        self.process_button.pack(side=tk.LEFT, padx=10)

        # 新增：生成报告按钮
        self.report_button = tk.Button(btn_container, text="生成报告",
                                       command=self.generate_report,
                                       font=('微软雅黑', 12, 'bold'),
                                       bg='#e67e22', fg='white',
                                       width=15, height=2)
        self.report_button.pack(side=tk.LEFT, padx=10)

        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=10)

        # 状态标签
        self.status_label = tk.Label(main_frame, text="准备就绪",
                                     font=('微软雅黑', 10),
                                     bg='#ecf0f1', fg='#7f8c8d')
        self.status_label.pack()

        # 日志文本框
        self.log_text = tk.Text(main_frame, height=8, width=70,
                                font=("Consolas", 9))
        self.log_text.pack(pady=10)  # 修复：添加pack布局

        # 滚动条
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical",
                                  command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)  # 修复：添加pack布局
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def select_input_file(self):
        """选择输入的ZIP文件"""
        filename = filedialog.askopenfilename(
            title="选择ZIP压缩包",
            filetypes=[("ZIP文件", "*.zip"), ("所有文件", "*.*")]
        )
        if filename:
            self.input_zip_path.set(filename)
            self.status_label.config(text=f"已选择: {os.path.basename(filename)}")
            self.log_message(f"已选择输入文件: {os.path.basename(filename)}")

    def select_output_file(self):
        """选择输出保存路径"""
        file_path = filedialog.asksaveasfilename(
            title="选择输出Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.output_file = file_path
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file_path)
            self.log_message(f"已选择输出文件: {os.path.basename(file_path)}")

    def decode_filename_gbk(self, filename):
        """使用GBK编码解码文件名"""
        try:
            return filename.encode('cp437').decode('gbk')
        except:
            return filename

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def run_processing(self, processing):
        if not self.input_file:
            messagebox.showerror("错误", "请选择输入文件！")
            return

        if not self.output_file:
            messagebox.showerror("错误", "请选择输出文件！")
            return

        if processing:
            # 开始处理时的状态
            self.process_button.config(state=tk.DISABLED, text="处理中...")
            self.progress.start()
            self.status_label.config(text="正在处理数据...")
        else:
            # 处理完成时的状态
            self.process_button.config(state=tk.NORMAL, text="开始处理")
            self.progress.stop()
            self.status_label.config(text="准备就绪")

        # 在新线程中运行处理函数，避免界面卡顿
        thread = threading.Thread(target=self.process_data)
        thread.daemon = True
        thread.start()

    def transform_rating_data(self, input_file):
        """
        分钟曲线df12
        将XLSX格式的收视率数据转换为每节目一行的结构

        参数:
            input_file: 输入XLSX文件路径
        """
        try:
            # 读取XLSX文件，列名为"名称"和"收视率%"
            df = input_file.iloc[2:].copy()
            df.columns = ['名称', '收视率%']

            # 检查必要的列是否存在
            if '名称' not in df.columns or '收视率%' not in df.columns:
                raise ValueError("输入文件必须包含'名称'和'收视率%'列")

            # 初始化变量
            programs = []
            current_program = None
            program_data = []
            program_start_time = None
            skip_rows = 0  # 跳过前30行

            # 遍历每一行数据
            for index, row in df.iterrows():
                if index < skip_rows:
                    continue

                name = row['名称']

                # 检查是否是时间数据（以"    << "开头）
                if isinstance(name, str) and name.startswith('    << '):
                    # 处理时间数据格式："    << 22:47 >>"
                    time_str = name.replace('    << ', '').replace(' >>', '').strip()
                    try:
                        # 自定义时间解析，支持24:00及以后的时间
                        hour, minute = map(int, time_str.split(':'))
                        if hour >= 24:
                            hour = hour - 24
                            # 创建第二天的日期对象
                            time_obj = datetime(2000, 1, 2, hour, minute)
                        else:
                            time_obj = datetime(2000, 1, 1, hour, minute)

                        if program_start_time is None:
                            program_start_time = time_obj

                        # 计算相对分钟数
                        relative_minute = int((time_obj - program_start_time).total_seconds() / 60)

                        # 处理收视率数据，使用Decimal实现精确四舍五入
                        rating_value = Decimal(str(row['收视率%']))
                        rating = float(rating_value.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))

                        # 添加到当前节目数据
                        program_data.append({
                            'minute': relative_minute,
                            'rating': rating
                        })

                    except ValueError:
                        continue
                else:
                    # 保存上一个节目数据
                    if current_program is not None and program_data:
                        programs.append({
                            'program': current_program,
                            'start_time': program_start_time,
                            'data': program_data
                        })

                    # 开始新节目
                    current_program = name
                    program_data = []
                    program_start_time = None

            # 添加最后一个节目
            if current_program is not None and program_data:
                programs.append({
                    'program': current_program,
                    'start_time': program_start_time,
                    'data': program_data
                })

            # 转换为DataFrame
            result_data = []
            for program in programs:
                # 按分钟排序
                sorted_data = sorted(program['data'], key=lambda x: x['minute'])

                # 创建每分钟一列的结构
                max_minute = max([d['minute'] for d in sorted_data]) if sorted_data else 0
                row = {'program': program['program']}

                # 填充每分钟的收视率
                for minute in range(max_minute + 1):
                    rating = next((d['rating'] for d in sorted_data if d['minute'] == minute), None)
                    row[f'{minute + 1}分钟'] = rating

                result_data.append(row)

            # 保存结果
            result_df = pd.DataFrame(result_data)
            return result_df

        except Exception as e:
            print(f"处理数据时出错: {str(e)}")
            raise

    # 4. 将开始时间和结束时间从"时:分:秒"格式改为"hh:mm"格式
    def format_time(self, time_str):
        try:
            # 尝试解析时间字符串
            time_obj = datetime.strptime(time_str, '%H:%M:%S')
            seconds = time_obj.second
            # 步骤4: 根据秒数决定是否进位
            if seconds >= 30:
                # 加30秒，让datetime自动处理进位
                rounded_time = time_obj + timedelta(seconds=30)
            else:
                # 不需要进位，直接减去秒数
                rounded_time = time_obj - timedelta(seconds=seconds)
            return rounded_time.strftime('%H:%M')
        except:
            # 如果格式不匹配，返回原值
            return time_str

    # 计算全国排名
    def calculate_national_rank(self, row, channel_name):
        # 获取所有频道收视率列
        rating_columns = [col for col in row.index if
                          col not in ['日期', '节目名称', '开始时间', '结束时间', '节目时长', '全国排名', '排名1',
                                      '排名2', '排名3', '排名变化']]

        # 获取该行的所有收视率值
        ratings = row[rating_columns].values

        # 获取中央电视台综合频道的收视率
        target_rating = row[channel_name]

        # 计算排名（降序排列，所以排名是1为最高）
        sorted_ratings = sorted(ratings, reverse=True)
        rank = 1
        for r in sorted_ratings:
            if r > target_rating:
                rank += 1
            else:
                break

        return rank

    # 获取排名前3的频道
    def get_top_channels(self, row):
        rating_columns = [col for col in row.index if
                          col not in ['日期', '节目名称', '开始时间', '结束时间', '节目时长', '全国排名', '排名1',
                                      '排名2', '排名3', '排名变化']]

        # 创建频道和收视率的字典
        channel_ratings = {col: row[col] for col in rating_columns}

        # 按收视率降序排序
        sorted_channels = sorted(channel_ratings.items(), key=lambda x: x[1], reverse=True)

        # 获取前3名
        top3 = [channel for channel, rating in sorted_channels[:3]]

        # 如果不足3个，用空字符串填充
        while len(top3) < 3:
            top3.append('')

        return top3

    # 字符替换函数（示例）
    def replace_chars(self, text, replacements):
        if pd.isna(text):
            return text
        for old, new in replacements.items():
            text = text.replace(old, new)
        return text

    def process_tv_data(self, input_file):
        """
        分钟总df11
        处理电视节目数据表

        参数:
            input_file: 输入的Excel文件路径
        """
        try:
            # 1. 从xlsx文件读取表格
            df = input_file.iloc[3:, 1:].copy()

            # 3. 重命名列
            new_columns = [
                '名称', '日期[相同值]', '周日[具体值]', '开始时间[最小]',
                '时长[总和]', '结束时间[最大]', '收视率%', '市场份额%', '平均忠实度'
            ]
            df.columns = new_columns

            # 4. 处理“周日[具体值]”列中的字符串
            df['周日[具体值]'] = df['周日[具体值]'].astype(str).str.replace('.', '').str.replace('七', '日')
            df['周日[具体值]'] = '周' + df['周日[具体值]']

            # 3. 收视率、市场份额、平均忠实度保留2位小数（严格四舍五入）
            numeric_columns = ['收视率%', '市场份额%', '平均忠实度']
            df[numeric_columns] = df[numeric_columns].map(
                lambda x: float(Decimal(str(x)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP)))

            # 时间格式化
            df['开始时间[最小]'] = df['开始时间[最小]'].apply(self.format_time)
            df['结束时间[最大]'] = df['结束时间[最大]'].apply(self.format_time)

            # 5. 按开始时间升序排列
            df = df.sort_values(by='开始时间[最小]')

            return df

        except Exception as e:
            print(f"处理过程中出现错误: {str(e)}")

    def process_share_data(self, input_file):
        """
        收视份额df2
        处理电视收视数据并生成分析结果

        参数:
            input_file: 输入的Excel文件路径
        """
        try:
            df = input_file.iloc[3:].copy()

            # 3. 重命名列
            new_columns = [
                '频道', '本日收视', '前日收视', '本日份额', '前日份额'
            ]
            df.columns = new_columns

            # 2. 按本日份额降序排列
            df = df.sort_values(by='本日份额', ascending=False)

            # 5. 新建"本日份额排名"列和"前日份额排名"列
            df['本日份额排名'] = df['本日份额'].rank(ascending=False, method='min').astype(int)
            df['前日份额排名'] = df['前日份额'].rank(ascending=False, method='min').astype(int)

            # 6. 新建"排名变化"列
            def calculate_rank_change(row):
                rank_change = row['前日份额排名'] - row['本日份额排名']

                if rank_change == 0:
                    return '持平'
                elif rank_change > 0:
                    return f'↑{abs(rank_change)}'
                else:
                    return f'↓{abs(rank_change)}'

            df['排名变化'] = df.apply(calculate_rank_change, axis=1)

            # 保留两位小数
            df['本日收视'] = df['本日收视'].apply(
                lambda x: Decimal(str(x)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))
            df['前日收视'] = df['前日收视'].apply(
                lambda x: Decimal(str(x)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))
            df['本日份额'] = df['本日份额'].apply(
                lambda x: Decimal(str(x)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))
            df['前日份额'] = df['前日份额'].apply(
                lambda x: Decimal(str(x)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))

            # 3. 新建"收视变化"列，按本日收视/前日收视-1计算
            df['收视变化'] = df.apply(lambda row:
                                      f"{Decimal(str((row['本日收视'] / row['前日收视'] - 1) * 100)).quantize(Decimal('1'), rounding=ROUND_HALF_UP)}%" if
                                      row['前日收视'] != 0 else 0, axis=1)

            # 4. 新建"份额变化"列，按本日份额/前日份额-1计算
            df['份额变化'] = df.apply(lambda row:
                                      f"{Decimal(str((row['本日份额'] / row['前日份额'] - 1) * 100)).quantize(Decimal('1'), rounding=ROUND_HALF_UP)}%" if
                                      row['前日份额'] != 0 else 0, axis=1)

            # 7. 调整表格列顺序
            columns_order = ['频道', '本日收视', '收视变化', '本日份额', '份额变化', '排名变化',
                             '前日收视', '前日份额', '本日份额排名', '前日份额排名']
            df = df[columns_order]

            # 排名列替换示例
            rank_replacements = {'中央电视台综合频道': '综合', '中央台八套': '电视剧', '中央台六套': '电影',
                                 '中央台五套': '体育', '中央台四套': '四套', '中央电视台新闻频道': '新闻',
                                 '湖南卫视': '湖南', '中央台三套': '综艺', '江苏卫视': '江苏', '浙江卫视': '浙江',
                                 '上海东方卫视': '东方', '北京卫视': '北京', '中央台二套': '财经', '广东卫视': '广东',
                                 '中央台七套': '军事', 'CCTV5+体育赛事频道': '体育赛事', '中央台九套纪录频道': '纪录',
                                 '中央电视台农业农村频道': '农业农村', '深圳卫视': '深圳', '山东卫视': '山东',
                                 '中央电视台少儿频道': '少儿', '湖南电视台金鹰卡通频道': '金鹰卡通', '安徽卫视': '安徽',
                                 '中央台十二套': '社会与法', '天津卫视': '天津', '中央台十一套': '戏曲',
                                 '中央电视台音乐频道': '音乐', '中央台十套': '科教', '卡酷少儿频道': '卡酷',
                                 '江西卫视': '江西', '湖北卫视': '湖北', '吉林卫视': '吉林', '辽宁卫视': '辽宁',
                                 '广东广播电视台嘉佳卡通频道': '嘉佳', '东南卫视': '东南', '河南卫视': '河南',
                                 '黑龙江卫视': '黑龙江', '贵州卫视': '贵州', '广西卫视': '广西', '青海卫视': '青海',
                                 '四川卫视': '四川', '内蒙古广播电视台内蒙古卫视频道': '内蒙古', '陕西卫视': '陕西',
                                 '西藏二套(汉语卫视)': '西藏汉语', '山西卫视': '山西', '甘肃卫视': '甘肃',
                                 '重庆卫视': '重庆', '宁夏卫视': '宁夏', '河北广播电视台卫视频道': '河北',
                                 '新疆卫视': '新疆', '凤凰卫视中文台': '凤凰中文', '海南卫视': '海南',
                                 '兵团卫视': '兵团', '优漫卡通卫视': '优漫', '云南广播电视台卫视频道(一套)': '云南',
                                 '厦门卫视': '厦门', 'CCTV4K超高清频道': '4K超高清', '山东教育卫视': '山东教育',
                                 '中国教育台一套': '教育一套', '凤凰卫视资讯台': '凤凰资讯',
                                 '新疆电视台二套(维语新闻综合频道)': '新疆维语新闻',
                                 '湖南电视台金鹰纪实频道': '金鹰纪实', '西藏一套(藏语卫视)': '西藏藏语', 'CGTN': 'CGTN',
                                 '新疆电视台三套(哈语新闻综合频道)': '新疆哈语', '福建海峡电视台': '福建海峡',
                                 '安多卫视': '安多', 'CCTV风云足球': '风云足球',
                                 '内蒙古广播电视台蒙古语卫视频道': '内蒙古-内蒙古语', '星空卫视': '星空',
                                 '香港卫视合家欢台': '合家欢', '阳光卫视': '阳光',
                                 '广东广播电视台大湾区卫视(上星版)': '大湾区卫视',
                                 '上海电视台纪实人文频道': '上海纪实人文', '凤凰卫视电影台': '凤凰电影',
                                 '上海电视台哈哈炫动频道': '哈哈炫动', '香港卫视音乐台': '香港卫视音乐',
                                 '香港卫视国际电影台': '香港卫视电影', '星空体育': '星空体育', '华娱卫视': '华娱',
                                 'MTV': 'MTV', '北京广播电视台体育休闲频道': '北京体育休闲',
                                 '香港有线卫视新知台': '香港卫视新知台', '农林卫视': '农林',
                                 '中国教育台二套': '教育二套', '参考:': '综合', }
            df['频道'] = df['频道'].apply(lambda x: self.replace_chars(x, rank_replacements))

            return df

        except Exception as e:
            print(f"处理数据时出错: {str(e)}")

    def process_channel_data(self, input_file, channel_name):
        # 一套变化情况df3 &电视剧频道黄金时段电视剧df4
        try:
            df = input_file.iloc[2:, 1:].copy()

            # 用第一行作为列名
            df.columns = df.iloc[0].copy()

            # 删除第一行（原列名行）
            df = df.iloc[1:].copy()
            df.reset_index(drop=True, inplace=True)

            # 删除摘要列和"参考:"列
            columns_to_drop = []
            for col in df.columns:
                if pd.notna(col) and ("摘要" in str(col) or "参考:" in str(col)):
                    columns_to_drop.append(col)

            df = df.drop(columns=columns_to_drop, errors='ignore')

            # 3. 重新命名列名
            column_mapping = {
                "日期 Tab": "日期",
                "名称/描述": "节目名称",
                "标题": "节目名称",
                "开始时间 Tab[最小]": "开始时间",
                "结束时间 Tab[具体值]": "结束时间",
                "结束时间 Tab[最大]": "结束时间",
                "时长[相同值]": "节目时长"
            }

            df = df.rename(columns=column_mapping)

            # 删除日期和节目名称列都为空的行
            if '日期' in df.columns and '节目名称' in df.columns:
                df = df.dropna(subset=['日期', '节目名称'], how='all')

            # 4. 处理日期列
            if '日期' in df.columns:
                # 前向填充日期列
                df['日期'] = df['日期'].ffill()

                # 获取所有唯一的日期值
                unique_dates = df['日期'].dropna().unique()

                if len(unique_dates) == 2:
                    # 转换为datetime进行比较
                    date_objects = []
                    for date_val in unique_dates:
                        if isinstance(date_val, str):
                            try:
                                date_obj = pd.to_datetime(date_val)
                                date_objects.append(date_obj)
                            except:
                                date_objects.append(date_val)
                        else:
                            date_objects.append(date_val)

                    # 找出较新和较旧的日期
                    if all(isinstance(d, pd.Timestamp) for d in date_objects):
                        newer_date = max(date_objects)
                        older_date = min(date_objects)

                        # 获取原始值
                        newer_original = unique_dates[date_objects.index(newer_date)]
                        older_original = unique_dates[date_objects.index(older_date)]

                        # 替换为"本日"和"前日"
                        df['日期'] = df['日期'].replace({newer_original: "本日", older_original: "前日"})

            # 5. 删除节目名称列为空的行
            if '节目名称' in df.columns:
                df = df.dropna(subset=['节目名称'])

            # 时间格式转换
            df['开始时间'] = df['开始时间'].apply(self.format_time)
            df['结束时间'] = df['结束时间'].apply(self.format_time)

            # 4. 排序：先按日期（本日在前，前日在后），再按开始时间升序
            df['日期排序'] = df['日期'].map({'本日': 0, '前日': 1})
            df = df.sort_values(['日期排序', '开始时间'])
            df = df.drop('日期排序', axis=1)

            # 5. 计算全国排名
            df['全国排名'] = df.apply(lambda row: self.calculate_national_rank(row, channel_name), axis=1)

            # 6. 获取排名前3的频道
            top_channels = df.apply(self.get_top_channels, axis=1, result_type='expand')
            df[['排名1', '排名2', '排名3']] = top_channels

            # 保留两位小数
            df[channel_name] = df[channel_name].apply(
                lambda x: Decimal(str(x)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))

            # 7. 按日期拆分表格
            today_df = df[df['日期'] == '本日'].copy()
            yesterday_df = df[df['日期'] == '前日'].copy()

            # 构建昨日数据映射
            yesterday_map = yesterday_df.set_index('节目名称').apply(
                lambda row: (row['开始时间'], row[channel_name]), axis=1
            ).to_dict()

            # 计算收视变化
            def calculate_change_with_time(row):
                yesterday_info = yesterday_map.get(row['节目名称'])
                if not yesterday_info:
                    return '/'
                yesterday_time_str, yesterday_value = yesterday_info
                today_time_str = row['开始时间']

                if yesterday_value is None or yesterday_value == 0:
                    return '/'

                try:
                    today_dt = datetime.strptime(today_time_str, '%H:%M')
                    yesterday_dt = datetime.strptime(yesterday_time_str, '%H:%M')
                except (ValueError, TypeError):
                    return '/'

                time_diff_seconds = abs((today_dt - yesterday_dt).total_seconds())
                if time_diff_seconds > 11700:
                    return '/'

                return f"{Decimal(str((row[channel_name] / yesterday_value - 1))).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP) * 100:.0f}%"

            today_df['收视变化'] = today_df.apply(calculate_change_with_time, axis=1)

            # 计算排名变化
            yesterday_rank_map = yesterday_df.set_index('节目名称').apply(
                lambda row: (row['开始时间'], row['全国排名']), axis=1).to_dict()

            def get_rank_change_with_time(row):
                program_name = row['节目名称']
                today_time_str = row['开始时间']
                today_rank = row['全国排名']

                yesterday_info = yesterday_rank_map.get(program_name)
                if not yesterday_info:
                    return '/'

                yesterday_time_str, yesterday_rank = yesterday_info
                if yesterday_rank is None:
                    return '/'

                try:
                    today_dt = datetime.strptime(today_time_str, '%H:%M')
                    yesterday_dt = datetime.strptime(yesterday_time_str, '%H:%M')
                except (ValueError, TypeError):
                    return '/'

                time_diff_seconds = abs((today_dt - yesterday_dt).total_seconds())
                if time_diff_seconds > 11700:
                    return '/'

                change = yesterday_rank - today_rank
                if change == 0:
                    return '持平'
                elif change > 0:
                    return f'↑{change}'
                else:
                    return f'↓{abs(change)}'

            today_df['排名变化'] = today_df.apply(get_rank_change_with_time, axis=1)

            # 字符替换
            program_replacements = {'朝闻天下': '朝闻TX', '生活圈': '生活Q', '无所畏惧之永': '无所畏惧之永',
                                    '新闻30分': 'XW30分', '今日说法': 'JR说法', '人生之路': '人生之路',
                                    '第1动画乐园': '第1动画LY', '农耕探文明': '农耕探文明', '新闻联播': 'XW联播',
                                    '焦点访谈': '焦点FT', '乌蒙深处': '乌蒙深处', '百年守护': '百年守护',
                                    '中华考工记': '中华考工记', '晚间新闻': '晚间XW', }
            today_df['节目名称'] = today_df['节目名称'].apply(lambda x: self.replace_chars(x, program_replacements))

            # 排名列替换
            rank_replacements = {'中央电视台综合频道': '综合', '中央台八套': '电视剧', '中央台六套': '电影',
                                 '中央台五套': '体育', '中央台四套': '四套', '中央电视台新闻频道': '新闻',
                                 '湖南卫视': '湖南', '中央台三套': '综艺', '江苏卫视': '江苏', '浙江卫视': '浙江',
                                 '上海东方卫视': '东方', '北京卫视': '北京', '中央台二套': '财经', '广东卫视': '广东',
                                 '中央台七套': '军事', 'CCTV5+体育赛事频道': '体育赛事', '中央台九套纪录频道': '纪录',
                                 '中央电视台农业农村频道': '农业农村', '深圳卫视': '深圳', '山东卫视': '山东',
                                 '中央电视台少儿频道': '少儿', '湖南电视台金鹰卡通频道': '金鹰卡通', '安徽卫视': '安徽',
                                 '中央台十二套': '社会与法', '天津卫视': '天津', '中央台十一套': '戏曲',
                                 '中央电视台音乐频道': '音乐', '中央台十套': '科教', '卡酷少儿频道': '卡酷',
                                 '江西卫视': '江西', '湖北卫视': '湖北', '吉林卫视': '吉林', '辽宁卫视': '辽宁',
                                 '广东广播电视台嘉佳卡通频道': '嘉佳', '东南卫视': '东南', '河南卫视': '河南',
                                 '黑龙江卫视': '黑龙江', '贵州卫视': '贵州', '广西卫视': '广西', '青海卫视': '青海',
                                 '四川卫视': '四川', '内蒙古广播电视台内蒙古卫视频道': '内蒙古', '陕西卫视': '陕西',
                                 '西藏二套(汉语卫视)': '西藏汉语', '山西卫视': '山西', '甘肃卫视': '甘肃',
                                 '重庆卫视': '重庆', '宁夏卫视': '宁夏', '河北广播电视台卫视频道': '河北',
                                 '新疆卫视': '新疆', '凤凰卫视中文台': '凤凰中文', '海南卫视': '海南',
                                 '兵团卫视': '兵团', '优漫卡通卫视': '优漫', '云南广播电视台卫视频道(一套)': '云南',
                                 '厦门卫视': '厦门', 'CCTV4K超高清频道': '4K超高清', '山东教育卫视': '山东教育',
                                 '中国教育台一套': '教育一套', '凤凰卫视资讯台': '凤凰资讯',
                                 '新疆电视台二套(维语新闻综合频道)': '新疆维语新闻',
                                 '湖南电视台金鹰纪实频道': '金鹰纪实', '西藏一套(藏语卫视)': '西藏藏语', 'CGTN': 'CGTN',
                                 '新疆电视台三套(哈语新闻综合频道)': '新疆哈语', '福建海峡电视台': '福建海峡',
                                 '安多卫视': '安多', 'CCTV风云足球': '风云足球',
                                 '内蒙古广播电视台蒙古语卫视频道': '内蒙古-内蒙古语', '星空卫视': '星空',
                                 '香港卫视合家欢台': '合家欢', '阳光卫视': '阳光',
                                 '广东广播电视台大湾区卫视(上星版)': '大湾区卫视',
                                 '上海电视台纪实人文频道': '上海纪实人文', '凤凰卫视电影台': '凤凰电影',
                                 '上海电视台哈哈炫动频道': '哈哈炫动', '香港卫视音乐台': '香港卫视音乐',
                                 '香港卫视国际电影台': '香港卫视电影', '星空体育': '星空体育', '华娱卫视': '华娱',
                                 'MTV': 'MTV', '北京广播电视台体育休闲频道': '北京体育休闲',
                                 '香港有线卫视新知台': '香港卫视新知台', '农林卫视': '农林',
                                 '中国教育台二套': '教育二套', '参考:': '综合', }
            today_df['排名1'] = today_df['排名1'].apply(lambda x: self.replace_chars(x, rank_replacements))
            today_df['排名2'] = today_df['排名2'].apply(lambda x: self.replace_chars(x, rank_replacements))
            today_df['排名3'] = today_df['排名3'].apply(lambda x: self.replace_chars(x, rank_replacements))

            # 整理表格列顺序
            columns_order = [
                                '开始时间', '节目名称', channel_name, '收视变化', '全国排名', '排名变化', '排名1',
                                '排名2', '排名3'
                            ] + [col for col in today_df.columns if
                                 col not in ['开始时间', '节目名称', '全国排名', '排名变化', '排名1', '排名2', '排名3']]

            today_df = today_df[columns_order]
            yesterday_df['收视变化'] = ''
            yesterday_df['排名变化'] = ''
            yesterday_df = yesterday_df[columns_order]

            final_df = pd.concat([today_df, yesterday_df], ignore_index=True)

            return final_df

        except Exception as e:
            print(f"处理数据时出错: {str(e)}")

    def read_specific_sheets_from_zip(self, zip_path):
        """
        从ZIP文件中读取指定的sheet数据
        """
        dataframes = {}

        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                file_list = zip_ref.namelist()

                # 解码文件名
                decoded_files = [self.decode_filename_gbk(f) for f in file_list]

                self.status_label.config(text="正在读取数据...")

                # 读取栏目分钟曲线文件
                for orig_file, decoded_file in zip(file_list, decoded_files):
                    if "栏目分钟曲线" in decoded_file and decoded_file.endswith('.xlsx'):
                        try:
                            with zip_ref.open(orig_file) as file:
                                xlsx_data = BytesIO(file.read())

                                # 读取交互分析sheet
                                df11 = pd.read_excel(xlsx_data, sheet_name="交互分析")
                                dataframes['dataframe11'] = df11
                                self.log_message(f"读取栏目分钟曲线-交互分析: {len(df11)} 行")

                                # 重新读取文件（因为pd.read_excel会消耗文件指针）
                                xlsx_data.seek(0)
                                # 读取每分钟节目sheet
                                df12 = pd.read_excel(xlsx_data, sheet_name="每分钟节目")
                                dataframes['dataframe12'] = df12
                                self.log_message(f"读取栏目分钟曲线-每分钟节目: {len(df12)} 行")

                        except Exception as e:
                            self.log_message(f"读取栏目分钟曲线失败: {str(e)}")

                    # 读取收视份额排名文件
                    elif "收视份额排名" in decoded_file and decoded_file.endswith('.xlsx'):
                        try:
                            with zip_ref.open(orig_file) as file:
                                xlsx_data = BytesIO(file.read())
                                df2 = pd.read_excel(xlsx_data, sheet_name="时期")
                                dataframes['dataframe2'] = df2
                                self.log_message(f"读取收视份额排名-时期: {len(df2)} 行")
                        except Exception as e:
                            self.log_message(f"读取收视份额排名失败: {str(e)}")

                    # 读取频道竞争分析（一套）文件
                    elif "频道竞争分析（一套）" in decoded_file and decoded_file.endswith('.xlsx'):
                        try:
                            with zip_ref.open(orig_file) as file:
                                xlsx_data = BytesIO(file.read())
                                df3 = pd.read_excel(xlsx_data, sheet_name="交互分析 (竞争）")
                                dataframes['dataframe3'] = df3
                                self.log_message(f"读取频道竞争分析（一套）: {len(df3)} 行")
                        except Exception as e:
                            self.log_message(f"读取频道竞争分析（一套）失败: {str(e)}")

                    # 读取频道竞争分析（八套）文件
                    elif "频道竞争分析（八套）" in decoded_file and decoded_file.endswith('.xlsx'):
                        try:
                            with zip_ref.open(orig_file) as file:
                                xlsx_data = BytesIO(file.read())
                                df4 = pd.read_excel(xlsx_data, sheet_name="交互分析 (竞争）")
                                dataframes['dataframe4'] = df4
                                self.log_message(f"读取频道竞争分析（八套）: {len(df4)} 行")
                        except Exception as e:
                            self.log_message(f"读取频道竞争分析（八套）失败: {str(e)}")

        except Exception as e:
            raise Exception(f"读取ZIP文件失败: {str(e)}")

        return dataframes

    def process_data(self):
        """处理数据的主函数"""
        try:
            # 检查输入路径
            if not self.input_zip_path.get():
                self.root.after(0, lambda: messagebox.showerror("错误", "请选择输入的ZIP文件"))
                return

            if not self.output_file:
                self.root.after(0, lambda: messagebox.showerror("错误", "请选择输出保存路径"))
                return

            # 读取数据
            self.log_message("正在读取ZIP文件...")
            dataframes = self.read_specific_sheets_from_zip(self.input_zip_path.get())

            # 检查是否成功读取所有数据
            expected_dfs = ['dataframe11', 'dataframe12', 'dataframe2', 'dataframe3', 'dataframe4']
            missing_dfs = [df for df in expected_dfs if df not in dataframes]

            if missing_dfs:
                self.root.after(0, lambda: messagebox.showwarning("警告", f"以下数据未找到: {', '.join(missing_dfs)}"))

            # 处理每个DataFrame
            self.log_message("正在处理【分钟曲线】...")
            dfa = self.transform_rating_data(dataframes['dataframe12'])
            dfa = dfa.rename (columns={'program': '名称'})

            self.log_message("正在处理【分钟表总收视】...")
            dfb = self.process_tv_data(dataframes['dataframe11'])

            self.log_message("正在处理【份额全国排名】...")
            dfc = self.process_share_data(dataframes['dataframe2'])

            self.log_message("正在处理【一套变化情况】...")
            dfd = self.process_channel_data(dataframes['dataframe3'], "中央电视台综合频道")

            self.log_message("正在处理【电视剧频道黄金时段电视剧】...")
            dfe = self.process_channel_data(dataframes['dataframe4'], "中央台八套")

            #节目名替换为全称
            # 定义不需要替换的关键词列表
            exclude_keywords = ['朝闻天下', '生活圈', '新闻30分', '今日说法', '第1动画乐园', '新闻联播', '焦点访谈', '晚间新闻']
            # 第一步：构建替换字典
            # Key: 全称的前6个字符
            # Value: 全称
            # 规则：如果全称中包含 exclude_keywords 中的任何一个，则不放入字典（即不替换）
            replace_map = {}

            for full_name in dfa['名称']:
                # 获取匹配键（前6个字符）
                key = full_name[:6]

                # 检查是否需要排除
                # 只要全称中包含任何一个关键词，就视为排除
                is_excluded = any(keyword in full_name for keyword in exclude_keywords)

                # 如果不在排除列表中，则建立映射关系
                if not is_excluded:
                    replace_map[key] = full_name

            print("\n--- 生成的替换映射 ---")
            for k, v in replace_map.items():
                print(f"简称前6位: '{k}' -> 全称: '{v}'")


            # 第二步：定义替换函数并应用到 dfb 和 dfd
            def replace_with_full_name(short_name):
                if pd.isna(short_name):
                    return short_name
                # 获取当前值的前6位作为查找键
                key = short_name[:6]
                # 在字典中查找，找到则返回全称，找不到返回原值
                return replace_map.get(key, short_name)

            # 对 dfb 进行替换
            dfb['名称'] = dfb['名称'].apply(replace_with_full_name)

            # 对 dfd 进行替换
            dfd['节目名称'] = dfd['节目名称'].apply(replace_with_full_name)

            # 保存结果
            self.log_message("正在保存结果...")

            # 创建ExcelWriter对象
            with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
                # 保存处理后的数据到不同的sheet
                dfa.to_excel(writer, sheet_name='分钟曲线', index=False)
                dfb.to_excel(writer, sheet_name='分钟总收视', index=False)
                dfc.to_excel(writer, sheet_name='份额全国排名', index=False)
                dfd.to_excel(writer, sheet_name='一套变化情况', index=False)
                dfe.to_excel(writer, sheet_name='电视剧频道黄金时段电视剧', index=False)

            self.log_message("数据处理完成！")

            # 在主线程中更新UI
            self.root.after(0, self.processing_complete, True)

        except Exception as e:
            error_msg = f"处理过程中发生错误: {str(e)}"
            self.log_message(error_msg)
            self.root.after(0, self.processing_complete, False, error_msg)

    def processing_complete(self, success, error_msg=None):
        self.progress.stop()
        self.process_button.config(state=tk.NORMAL)

        if success:
            self.status_label.config(text="处理完成！")
            self.root.after(0, lambda: messagebox.showinfo("成功",
                                                           f"数据处理完成！\n输出文件已保存至：\n{self.output_file}"))
        else:
            self.status_label.config(text="处理失败！")
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))

    def start_processing(self):
        """开始处理（在新线程中运行）"""
        # 禁用按钮，防止重复点击
        self.process_button.config(state=tk.DISABLED)

        # 启动进度条
        self.progress.start()

        # 在新线程中运行处理函数，避免界面卡死
        processing_thread = threading.Thread(target=self.process_data)
        processing_thread.daemon = True
        processing_thread.start()

    # 修改generate_report方法
    def generate_report(self):
        def run_report():
            try:
                # 检查输出文件是否存在
                if not self.output_file:
                    self.root.after(0, lambda: messagebox.showerror("错误", "请先选择输出文件路径并完成数据处理！"))
                    return

                if not os.path.exists(self.output_file):
                    self.root.after(0, lambda: messagebox.showerror("错误", f"输出文件不存在：{self.output_file}\n请先点击'处理数据'生成文件！"))
                    return

                # 禁用报告按钮
                self.root.after(0, lambda: self.report_button.config(state=tk.DISABLED, text="生成中..."))
                self.root.after(0, lambda: self.status_label.config(text="正在生成报告..."))
                self.log_message("开始生成报告...")

                # 配置文件路径
                current_dir = os.path.dirname(os.path.abspath(__file__))
                # 拼接res目录下的模板文件路径
                word_template = os.path.join(current_dir, "res", "报告模板.docx")
                # 检查模板文件是否存在
                if not os.path.exists(word_template):
                    raise FileNotFoundError(f"报告模板文件不存在：{word_template}\n请确保该文件在程序目录下")

                # 创建报告生成器实例
                report_generator = sub.ReportGenerator()
                # 对接日志系统
                report_generator.set_logger(self.log_message)

                # 调用生成报告方法
                success, msg, output_file = report_generator.generate_report(
                    excel_path=self.output_file,
                    word_template_path=word_template
                )

                # 更新UI状态
                if success:
                    self.log_message(f"报告生成成功！文件路径：{output_file}")
                    self.root.after(0, lambda: messagebox.showinfo("成功", f"报告生成完成！\n文件保存至：\n{output_file}"))
                    self.root.after(0, lambda: self.status_label.config(text="报告生成完成！"))
                else:
                    self.log_message(f"报告生成失败：{msg}")
                    self.root.after(0, lambda: messagebox.showerror("错误", f"生成报告失败：\n{msg}"))
                    self.root.after(0, lambda: self.status_label.config(text="报告生成失败！"))

                # 调用生成报告方法
                success, msg, output_file = report_generator.merge_tv_ratings_data(
                    input_file=self.output_file
                )
                # 更新UI状态
                if success:
                    self.log_message(f"合并曲线生成成功！文件路径：{output_file}")
                    self.root.after(0, lambda: messagebox.showinfo("成功",
                                                                   f"合并曲线生成完成！\n文件保存至：\n{output_file}"))
                    self.root.after(0, lambda: self.status_label.config(text="合并曲线生成完成！"))
                else:
                    self.log_message(f"合并曲线生成失败：{msg}")
                    self.root.after(0, lambda: messagebox.showerror("错误", f"生成合并曲线失败：\n{msg}"))
                    self.root.after(0, lambda: self.status_label.config(text="合并曲线生成失败！"))

            except Exception as e:
                error_msg = f"生成报告时发生错误: {str(e)}"
                self.log_message(error_msg)
                self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
                self.root.after(0, lambda: self.status_label.config(text="报告生成失败！"))
            finally:
                # 恢复按钮状态
                self.root.after(0, lambda: self.report_button.config(state=tk.NORMAL, text="生成报告"))


        # 在新线程中执行，避免界面卡顿
        report_thread = threading.Thread(target=run_report)
        report_thread.daemon = True
        report_thread.start()

def main():
    """主函数"""
    root = tk.Tk()
    app = DataProcessorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
