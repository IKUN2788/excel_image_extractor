#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel图片提取工具
功能：从Excel文件中提取浮动组合图片，按单元格地址创建目录保存
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import zipfile
import shutil
from datetime import datetime
import xml.etree.ElementTree as ET
import re
from pathlib import Path
import threading
import hashlib
try:
    from PIL import Image
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False


class ExcelImageExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel图片提取工具")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 变量
        self.selected_file = tk.StringVar()
        self.output_dir = tk.StringVar()
        
        # 设置默认输出目录为当前目录
        self.output_dir.set(os.getcwd())
        
        self.setup_ui()
        
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="Excel图片提取工具", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="Excel文件:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(file_frame, textvariable=self.selected_file, state="readonly").grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(file_frame, text="浏览", command=self.browse_file).grid(row=0, column=2)
        
        # 输出目录选择
        ttk.Label(file_frame, text="输出目录:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        ttk.Entry(file_frame, textvariable=self.output_dir, state="readonly").grid(
            row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(10, 0))
        ttk.Button(file_frame, text="选择", command=self.browse_output_dir).grid(row=1, column=2, pady=(10, 0))
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        self.extract_button = ttk.Button(button_frame, text="开始提取", 
                                        command=self.start_extraction, state="disabled")
        self.extract_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="清空日志", command=self.clear_log).pack(side=tk.LEFT)
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="提取日志", padding="10")
        log_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, state="disabled")
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def browse_file(self):
        """浏览选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.selected_file.set(file_path)
            self.extract_button.config(state="normal")
            self.log_message(f"已选择文件: {file_path}")
            
    def browse_output_dir(self):
        """浏览选择输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_dir.set(dir_path)
            self.log_message(f"输出目录设置为: {dir_path}")
            
    def log_message(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        
        # 更新界面
        self.root.update_idletasks()
        
    def clear_log(self):
        """清空日志"""
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")
        
    def calculate_file_hash(self, file_path):
        """计算文件的MD5哈希值"""
        try:
            hash_md5 = hashlib.md5()
            with open(file_path, "rb") as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_md5.update(chunk)
            return hash_md5.hexdigest()
        except Exception as e:
            self.log_message(f"计算文件哈希时出错 {file_path}: {str(e)}")
            return None
            
    def get_unique_filename(self, base_path, filename, hash_value, hash_tracker):
        """生成唯一的文件名，处理重复图片"""
        name, ext = os.path.splitext(filename)
        
        # 检查是否是重复图片
        if hash_value in hash_tracker:
            # 这是重复图片，添加序号
            hash_tracker[hash_value]['count'] += 1
            count = hash_tracker[hash_value]['count']
            unique_filename = f"{name}_副本{count}{ext}"
            self.log_message(f"检测到重复图片: {filename} -> {unique_filename} (哈希: {hash_value[:8]}...)")
        else:
            # 这是新图片
            hash_tracker[hash_value] = {'count': 0, 'original_name': filename}
            unique_filename = filename
            
        # 确保文件名在目标目录中是唯一的
        counter = 1
        final_filename = unique_filename
        while os.path.exists(os.path.join(base_path, final_filename)):
            name_part, ext_part = os.path.splitext(unique_filename)
            final_filename = f"{name_part}_{counter}{ext_part}"
            counter += 1
            
        return final_filename
        
    def start_extraction(self):
        """开始提取图片（在新线程中运行）"""
        if not self.selected_file.get():
            messagebox.showerror("错误", "请先选择Excel文件")
            return
            
        # 在新线程中运行提取过程，避免界面冻结
        thread = threading.Thread(target=self.extract_images)
        thread.daemon = True
        thread.start()
        
    def extract_images(self):
        """提取Excel文件中的图片"""
        try:
            self.progress.start()
            self.extract_button.config(state="disabled")
            self.status_var.set("正在提取图片...")
            
            excel_file = self.selected_file.get()
            output_base = self.output_dir.get()
            
            self.log_message("开始分析Excel文件...")
            
            # 创建临时目录
            temp_dir = os.path.join(output_base, "temp_excel_extract")
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)
            
            try:
                # 解压Excel文件（XLSX实际上是ZIP文件）
                with zipfile.ZipFile(excel_file, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                self.log_message("Excel文件解压完成")
                
                # 查找图片文件
                media_dir = os.path.join(temp_dir, "xl", "media")
                drawings_dir = os.path.join(temp_dir, "xl", "drawings")
                worksheets_dir = os.path.join(temp_dir, "xl", "worksheets")
                
                if not os.path.exists(media_dir):
                    self.log_message("未找到图片文件")
                    return
                
                # 获取所有图片文件
                image_files = []
                for file in os.listdir(media_dir):
                    if file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                        image_files.append(file)
                
                self.log_message(f"找到 {len(image_files)} 个图片文件")
                
                if not image_files:
                    self.log_message("没有找到图片文件")
                    return
                
                # 分析绘图关系和位置
                image_locations = self.analyze_image_locations(drawings_dir, worksheets_dir)
                
                # 创建主提取结果目录
                main_output_dir = os.path.join(output_base, "提取结果")
                os.makedirs(main_output_dir, exist_ok=True)
                self.log_message(f"创建主输出目录: {main_output_dir}")
                
                # 创建输出目录并复制图片
                extracted_count = 0
                processed_groups = {}  # 记录已处理的组合图片
                hash_tracker = {}  # 记录图片哈希值，用于检测重复
                duplicate_count = 0  # 重复图片计数
                
                self.log_message("开始检测重复图片...")
                
                for image_file in image_files:
                    # 获取图片位置信息列表
                    location_list = image_locations.get(image_file, [])
                    if not location_list:
                        # 如果没有位置信息，使用默认值
                        location_list = [{'cell': 'Unknown', 'sheet': 'Sheet1', 'is_group': False}]
                    
                    # 为每个位置创建目录并复制图片
                    for location_info in location_list:
                        cell_address = location_info.get('cell', 'Unknown')
                        sheet_name = location_info.get('sheet', 'Sheet1')
                        is_group = location_info.get('is_group', False)
                        
                        # 创建以单元格地址命名的目录（在"提取结果"目录下）
                        safe_cell_name = re.sub(r'[<>:"/\\|?*]', '_', f"{sheet_name}_{cell_address}")
                        output_dir = os.path.join(main_output_dir, safe_cell_name)
                        
                        # 如果是组合图片，确保所有图片都放在同一个目录
                        if is_group:
                            # 检查是否已经为这个位置创建了目录
                            group_key = f"{sheet_name}_{cell_address}"
                            if group_key not in processed_groups:
                                processed_groups[group_key] = []
                                self.log_message(f"创建组合图片目录: 提取结果/{safe_cell_name}")
                            
                            processed_groups[group_key].append(image_file)
                        
                        os.makedirs(output_dir, exist_ok=True)
                        
                        # 复制图片文件
                        src_path = os.path.join(media_dir, image_file)
                        if os.path.exists(src_path):
                            # 计算图片哈希值
                            file_hash = self.calculate_file_hash(src_path)
                            if file_hash:
                                # 生成唯一文件名
                                unique_filename = self.get_unique_filename(output_dir, image_file, file_hash, hash_tracker)
                                dst_path = os.path.join(output_dir, unique_filename)
                                
                                # 检查是否是重复图片
                                if unique_filename != image_file:
                                    duplicate_count += 1
                                
                                shutil.copy2(src_path, dst_path)
                                
                                extracted_count += 1
                                if is_group:
                                    self.log_message(f"已提取组合图片: {image_file} -> 提取结果/{safe_cell_name}/{unique_filename}")
                                else:
                                    self.log_message(f"已提取单独图片: {image_file} -> 提取结果/{safe_cell_name}/{unique_filename}")
                            else:
                                self.log_message(f"警告: 无法计算图片哈希值: {src_path}")
                        else:
                            self.log_message(f"警告: 图片文件不存在: {src_path}")
                
                # 输出组合图片统计信息
                for group_key, images in processed_groups.items():
                    if len(images) > 1:
                        self.log_message(f"组合图片 {group_key} 包含 {len(images)} 张图片: {', '.join(images)}")
                
                # 输出重复图片统计信息
                unique_images = len(hash_tracker)
                self.log_message(f"重复图片检测完成！")
                self.log_message(f"- 总图片数: {len(image_files)}")
                self.log_message(f"- 唯一图片数: {unique_images}")
                self.log_message(f"- 重复图片数: {duplicate_count}")
                if duplicate_count > 0:
                    self.log_message(f"- 重复图片已自动重命名，添加'_副本'后缀")
                
                self.log_message(f"提取完成！共提取 {extracted_count} 个图片文件到 '提取结果' 目录")
                
                # 执行图片合并
                if PILLOW_AVAILABLE:
                    self.log_message("开始合并图片...")
                    merged_count = self.merge_images(main_output_dir)
                    self.log_message(f"图片合并完成！共合并 {merged_count} 个目录的图片")
                    self.status_var.set(f"提取并合并完成，共 {extracted_count} 个文件，{merged_count} 个合并图片")
                    
                    messagebox.showinfo("完成", f"图片提取和合并完成！\n共提取 {extracted_count} 个图片文件\n组合图片组数: {len(processed_groups)}\n合并图片数: {merged_count}\n唯一图片数: {unique_images}\n重复图片数: {duplicate_count}\n\n所有文件已保存到 '提取结果' 和 '合并结果' 目录中")
                else:
                    self.log_message("警告: 未安装Pillow库，跳过图片合并功能")
                    self.status_var.set(f"提取完成，共 {extracted_count} 个文件")
                    messagebox.showinfo("完成", f"图片提取完成！\n共提取 {extracted_count} 个图片文件\n组合图片组数: {len(processed_groups)}\n唯一图片数: {unique_images}\n重复图片数: {duplicate_count}\n\n所有文件已保存到 '提取结果' 目录中\n\n注意: 未安装Pillow库，无法进行图片合并")
                
            finally:
                # 清理临时目录
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    
        except Exception as e:
            error_msg = f"提取过程中发生错误: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("错误", error_msg)
            self.status_var.set("提取失败")
            
        finally:
            self.progress.stop()
            self.extract_button.config(state="normal")
            
    def analyze_image_locations(self, drawings_dir, worksheets_dir):
        """分析图片在Excel中的位置"""
        image_locations = {}
        
        try:
            # 首先解析关系文件，建立图片ID到文件名的映射
            rels_mapping = self.parse_drawing_rels(drawings_dir)
            
            # 分析绘图文件
            if os.path.exists(drawings_dir):
                drawing_files = [f for f in os.listdir(drawings_dir) if f.endswith('.xml')]
                self.log_message(f"找到 {len(drawing_files)} 个绘图文件")
                
                for drawing_file in drawing_files:
                    drawing_path = os.path.join(drawings_dir, drawing_file)
                    self.parse_drawing_xml(drawing_path, image_locations, drawing_file, rels_mapping)
            
            # 分析工作表关系
            if os.path.exists(worksheets_dir):
                self.analyze_worksheet_relations(worksheets_dir, image_locations)
            
            # 如果没有找到位置信息，使用默认位置
            if not image_locations:
                self.log_message("未能确定图片具体位置，使用默认位置")
                # 为每个图片分配默认位置
                media_dir = os.path.dirname(drawings_dir).replace("drawings", "media")
                if os.path.exists(media_dir):
                    image_files = [f for f in os.listdir(media_dir) 
                                 if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
                    for i, image_file in enumerate(image_files):
                        col = i % 10  # 每行10个图片
                        row = i // 10 + 1
                        cell_address = self.col_num_to_letter(col) + str(row)
                        image_locations[image_file] = {
                            'cell': cell_address,
                            'sheet': 'Sheet1'
                        }
                
        except Exception as e:
            self.log_message(f"分析图片位置时出错: {str(e)}")
            
        return image_locations

    def parse_drawing_rels(self, drawings_dir):
        """解析绘图关系文件，建立图片ID到文件名的映射"""
        rels_mapping = {}
        
        try:
            rels_dir = os.path.join(drawings_dir, "_rels")
            if os.path.exists(rels_dir):
                for rels_file in os.listdir(rels_dir):
                    if rels_file.endswith('.xml.rels'):
                        rels_path = os.path.join(rels_dir, rels_file)
                        tree = ET.parse(rels_path)
                        root = tree.getroot()
                        
                        # 解析关系映射
                        for relationship in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            rel_id = relationship.get('Id')
                            target = relationship.get('Target')
                            rel_type = relationship.get('Type')
                            
                            # 只处理图片关系
                            if rel_type and 'image' in rel_type and target:
                                # 提取图片文件名
                                image_filename = os.path.basename(target)
                                rels_mapping[rel_id] = image_filename
                                self.log_message(f"关系映射: {rel_id} -> {image_filename}")
                        
                        # 为这个绘图文件保存映射
                        drawing_name = rels_file.replace('.xml.rels', '.xml')
                        if drawing_name not in rels_mapping:
                            rels_mapping[drawing_name] = {}
                        rels_mapping[drawing_name] = dict(rels_mapping)
                        
        except Exception as e:
            self.log_message(f"解析关系文件时出错: {str(e)}")
            
        return rels_mapping
        
    def analyze_worksheet_relations(self, worksheets_dir, image_locations):
        """分析工作表与绘图的关系"""
        try:
            for worksheet_file in os.listdir(worksheets_dir):
                if worksheet_file.endswith('.xml'):
                    worksheet_path = os.path.join(worksheets_dir, worksheet_file)
                    self.parse_worksheet_xml(worksheet_path, image_locations, worksheet_file)
        except Exception as e:
            self.log_message(f"分析工作表关系时出错: {str(e)}")
    
    def parse_worksheet_xml(self, xml_path, image_locations, worksheet_file):
        """解析工作表XML文件"""
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # 获取工作表名称
            sheet_name = worksheet_file.replace('.xml', '').replace('sheet', 'Sheet')
            
            # 查找绘图引用
            namespaces = {
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            }
            
            # 更新已有的图片位置信息，添加正确的工作表名称
            for image_file, location in image_locations.items():
                if location.get('sheet') == 'Sheet1':  # 默认工作表
                    location['sheet'] = sheet_name
                    
        except Exception as e:
            self.log_message(f"解析工作表XML文件 {xml_path} 时出错: {str(e)}")

    def parse_drawing_xml(self, xml_path, image_locations, drawing_file, rels_mapping):
        """解析绘图XML文件获取图片位置"""
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # Excel绘图XML的命名空间
            namespaces = {
                'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            }
            
            pic_count = 0
            group_images = {}  # 用于存储组合图片信息
            
            # 查找所有的图片元素和组合元素
            for anchor in root.findall('.//xdr:oneCellAnchor', namespaces) + root.findall('.//xdr:twoCellAnchor', namespaces):
                # 检查是否是组合图片
                grp_sp = anchor.find('.//xdr:grpSp', namespaces)
                if grp_sp is not None:
                    # 这是一个组合图片
                    self.log_message(f"发现组合图片在 {drawing_file}")
                    
                    # 获取组合的位置信息
                    from_elem = anchor.find('xdr:from', namespaces)
                    if from_elem is not None:
                        col_elem = from_elem.find('xdr:col', namespaces)
                        row_elem = from_elem.find('xdr:row', namespaces)
                        
                        if col_elem is not None and row_elem is not None:
                            try:
                                col = int(col_elem.text)
                                row = int(row_elem.text) + 1  # Excel行号从1开始
                                cell_address = self.col_num_to_letter(col) + str(row)
                                
                                # 查找组合中的所有图片
                                pics_in_group = grp_sp.findall('.//xdr:pic', namespaces)
                                self.log_message(f"组合图片位置: {cell_address}, 包含 {len(pics_in_group)} 张图片")
                                
                                for pic_elem in pics_in_group:
                                    pic_count += 1
                                    
                                    # 获取图片的关系ID
                                    blip_elem = pic_elem.find('.//a:blip', namespaces)
                                    if blip_elem is not None:
                                        embed_id = blip_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                        if embed_id and embed_id in rels_mapping:
                                            image_filename = rels_mapping[embed_id]
                                            # 支持同一图片文件在多个位置
                                            if image_filename not in image_locations:
                                                image_locations[image_filename] = []
                                            image_locations[image_filename].append({
                                                'cell': cell_address,
                                                'sheet': 'Sheet1',
                                                'embed_id': embed_id,
                                                'is_group': True,
                                                'group_position': pic_count
                                            })
                                            self.log_message(f"组合图片: {image_filename} -> {cell_address}")
                                        else:
                                            # 如果没有找到关系映射，使用默认命名
                                            image_filename = f"image{pic_count}.png"
                                            if image_filename not in image_locations:
                                                image_locations[image_filename] = []
                                            image_locations[image_filename].append({
                                                'cell': cell_address,
                                                'sheet': 'Sheet1',
                                                'embed_id': embed_id or f"rId{pic_count}",
                                                'is_group': True,
                                                'group_position': pic_count
                                            })
                                            self.log_message(f"组合图片(默认命名): {image_filename} -> {cell_address}")
                                            
                            except (ValueError, TypeError) as e:
                                self.log_message(f"解析组合图片位置坐标时出错: {str(e)}")
                else:
                    # 单独的图片
                    pic_elem = anchor.find('.//xdr:pic', namespaces)
                    if pic_elem is not None:
                        pic_count += 1
                        
                        # 获取位置信息
                        from_elem = anchor.find('xdr:from', namespaces)
                        if from_elem is not None:
                            col_elem = from_elem.find('xdr:col', namespaces)
                            row_elem = from_elem.find('xdr:row', namespaces)
                            
                            if col_elem is not None and row_elem is not None:
                                try:
                                    col = int(col_elem.text)
                                    row = int(row_elem.text) + 1  # Excel行号从1开始
                                    
                                    # 转换为Excel单元格地址
                                    cell_address = self.col_num_to_letter(col) + str(row)
                                    
                                    # 尝试获取图片的关系ID
                                    blip_elem = pic_elem.find('.//a:blip', namespaces)
                                    if blip_elem is not None:
                                        embed_id = blip_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                        if embed_id and embed_id in rels_mapping:
                                            image_filename = rels_mapping[embed_id]
                                            # 支持同一图片文件在多个位置
                                            if image_filename not in image_locations:
                                                image_locations[image_filename] = []
                                            image_locations[image_filename].append({
                                                'cell': cell_address,
                                                'sheet': 'Sheet1',
                                                'embed_id': embed_id,
                                                'is_group': False
                                            })
                                            self.log_message(f"单独图片: {image_filename} -> {cell_address}")
                                        else:
                                            # 如果没有找到关系映射，使用默认命名
                                            image_filename = f"image{pic_count}.png"
                                            if image_filename not in image_locations:
                                                image_locations[image_filename] = []
                                            image_locations[image_filename].append({
                                                'cell': cell_address,
                                                'sheet': 'Sheet1',
                                                'embed_id': embed_id or f"rId{pic_count}",
                                                'is_group': False
                                            })
                                            self.log_message(f"单独图片(默认命名): {image_filename} -> {cell_address}")
                                    
                                except (ValueError, TypeError) as e:
                                    self.log_message(f"解析位置坐标时出错: {str(e)}")
            
            if pic_count > 0:
                self.log_message(f"在 {drawing_file} 中找到 {pic_count} 个图片")
            else:
                self.log_message(f"在 {drawing_file} 中未找到图片")
                        
        except Exception as e:
            self.log_message(f"解析绘图XML文件 {xml_path} 时出错: {str(e)}")
            
    def col_num_to_letter(self, col_num):
        """将列号转换为Excel列字母"""
        result = ""
        while col_num >= 0:
            result = chr(col_num % 26 + ord('A')) + result
            col_num = col_num // 26 - 1
            if col_num < 0:
                break
        return result
    
    def merge_images(self, extraction_dir):
        """合并每个目录中的图片为一张横向排列的图片"""
        if not PILLOW_AVAILABLE:
            self.log_message("错误: 未安装Pillow库，无法进行图片合并")
            return 0
        
        merged_count = 0
        
        try:
            # 创建合并结果目录
            merge_output_dir = os.path.join(os.path.dirname(extraction_dir), "合并结果")
            os.makedirs(merge_output_dir, exist_ok=True)
            self.log_message(f"创建合并结果目录: {merge_output_dir}")
            
            # 遍历提取结果目录中的所有子目录
            for item in os.listdir(extraction_dir):
                item_path = os.path.join(extraction_dir, item)
                
                # 只处理目录
                if os.path.isdir(item_path):
                    # 获取目录中的所有图片文件
                    image_files = []
                    for file in os.listdir(item_path):
                        if file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                            image_files.append(os.path.join(item_path, file))
                    
                    # 如果目录中有图片文件，进行合并
                    if image_files:
                        # 按文件名排序，确保合并顺序一致
                        image_files.sort()
                        
                        if len(image_files) == 1:
                            # 只有一张图片，直接复制
                            src_file = image_files[0]
                            dst_file = os.path.join(merge_output_dir, f"{item}.png")
                            shutil.copy2(src_file, dst_file)
                            self.log_message(f"单张图片复制: {item}")
                        else:
                            # 多张图片，进行横向合并
                            merged_image = self.merge_images_horizontally(image_files)
                            if merged_image:
                                output_file = os.path.join(merge_output_dir, f"{item}_合并.png")
                                merged_image.save(output_file, "PNG")
                                self.log_message(f"合并完成: {item} ({len(image_files)} 张图片)")
                            else:
                                self.log_message(f"合并失败: {item}")
                                continue
                        
                        merged_count += 1
                    else:
                        self.log_message(f"目录 {item} 中没有图片文件")
            
        except Exception as e:
            self.log_message(f"图片合并过程中发生错误: {str(e)}")
        
        return merged_count
    
    def merge_images_horizontally(self, image_paths):
        """将多张图片横向合并为一张图片"""
        try:
            # 打开所有图片
            images = []
            for path in image_paths:
                try:
                    img = Image.open(path)
                    # 转换为RGBA模式以支持透明度
                    if img.mode != 'RGBA':
                        img = img.convert('RGBA')
                    images.append(img)
                except Exception as e:
                    self.log_message(f"无法打开图片 {path}: {str(e)}")
                    continue
            
            if not images:
                return None
            
            # 计算合并后图片的尺寸
            # 高度取所有图片的最大高度
            max_height = max(img.height for img in images)
            # 宽度为所有图片宽度之和
            total_width = sum(img.width for img in images)
            
            # 创建新的空白图片
            merged_image = Image.new('RGBA', (total_width, max_height), (255, 255, 255, 0))
            
            # 将图片依次粘贴到合并图片上
            x_offset = 0
            for img in images:
                # 计算垂直居中位置
                y_offset = (max_height - img.height) // 2
                merged_image.paste(img, (x_offset, y_offset), img)
                x_offset += img.width
            
            return merged_image
            
        except Exception as e:
            self.log_message(f"图片合并过程中发生错误: {str(e)}")
            return None


def main():
    """主函数"""
    root = tk.Tk()
    app = ExcelImageExtractor(root)
    
    # 设置窗口图标（如果有的话）
    try:
        # root.iconbitmap('icon.ico')  # 如果有图标文件
        pass
    except:
        pass
    
    # 居中显示窗口
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()