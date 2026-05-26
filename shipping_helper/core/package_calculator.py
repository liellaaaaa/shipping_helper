# -*- coding: utf-8 -*-
"""
包装计算器 - 根据订单要求和包装资料计算桶数、卡板数、毛重、净重、体积
参考 d:\code\船务部\calculate\calculate.py 的计算逻辑
"""

import os
import re
import json


class PackageCalculator:
    """包装计算器"""

    def __init__(self, packaging_file: str = None):
        self.packaging_file = packaging_file
        self.packages = []  # 桶类型列表
        self.pallets = []   # 卡板列表
        self.pallet_capacity = {}  # 每种桶对应的卡板容量，动态从JSON解析
        self.container_specs = {
            '20gp': {
                'length_m': 5.898,
                'width_m': 2.352,
                'height_m': 2.385,
                'door_height_m': 2.28,
                'door_width_m': 2.343,
                'max_cbm': 5.898 * 2.352 * 2.385,  # ≈33.07
                'max_pallets': 20,
                'pallets_per_layer': 10,
            },
            '40gp': {
                'length_m': 12.032,
                'width_m': 2.352,
                'height_m': 2.385,
                'door_height_m': 2.28,
                'door_width_m': 2.343,
                'max_cbm': 12.032 * 2.352 * 2.385,
                'max_pallets': 40,
                'pallets_per_layer': 20,
            },
        }
        self.no_pallet_container_capacity = {}  # 不打卡板时货柜容量 {"125kg": (116, "20'货柜可以装116桶"), ...}

        self._load_from_json(packaging_file)

    def _load_from_json(self, filepath: str):
        """从XLS或JSON文件加载包装资料"""
        if not filepath or not os.path.exists(filepath):
            # 尝试从data目录下查找
            data_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', '01_下单生产')
            if os.path.exists(data_dir):
                for f in os.listdir(data_dir):
                    if '产品包装' in f and (f.endswith('.xls') or f.endswith('.xlsx')):
                        filepath = os.path.join(data_dir, f)
                        break

        if not filepath or not os.path.exists(filepath):
            print(f"警告: 包装资料文件不存在: {filepath}")
            self._set_defaults()
            return

        try:
            # 根据文件扩展名选择加载方式
            if filepath.endswith('.xls') or filepath.endswith('.xlsx'):
                self._load_from_xls(filepath)
            else:
                self._load_from_json_file(filepath)
        except Exception as e:
            print(f"加载包装资料失败: {e}")
            self._set_defaults()

    def _load_from_xls(self, filepath: str):
        """从XLS/XLSX文件加载包装资料"""
        is_xlsx = filepath.endswith('.xlsx')

        if is_xlsx:
            # openpyxl只支持xlsx格式
            import openpyxl
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = wb.active
            self._parse_xls_worksheet_from_openpyxl(ws)
        else:
            # xlrd只支持旧格式xls
            import xlrd
            wb = xlrd.open_workbook(filepath, encoding_override='utf-8')
            ws = wb.sheet_by_index(0)
            self._parse_xls_worksheet(ws)

    def _parse_xls_worksheet(self, ws):
        """解析xlrd worksheet"""
        for i in range(2, ws.nrows):
            row = ws.row_values(i)
            self._process_package_row(row)

        if not self.pallets:
            self.pallets = [
                {'name': '1.0*1.0m卡板', 'size_m': '1.0*1.0', 'tare_kg': 17.0, 'cbm': 0.15},
                {'name': '1.1*1.1m卡板', 'size_m': '1.1*1.1', 'tare_kg': 18.5, 'cbm': 0.2},
            ]

        if not self.pallet_capacity:
            self._set_default_pallet_capacity()

    def _parse_xls_worksheet_from_openpyxl(self, ws):
        """解析openpyxl worksheet"""
        for i in range(3, ws.max_row + 1):
            row = [ws.cell(i, j).value for j in range(1, 6)]
            self._process_package_row(row)

        if not self.pallets:
            self.pallets = [
                {'name': '1.0*1.0m卡板', 'size_m': '1.0*1.0', 'tare_kg': 17.0, 'cbm': 0.15},
                {'name': '1.1*1.1m卡板', 'size_m': '1.1*1.1', 'tare_kg': 18.5, 'cbm': 0.2},
            ]

        if not self.pallet_capacity:
            self._set_default_pallet_capacity()

    def _process_package_row(self, row):
        """处理一行包装数据"""
        if not row[0] or not str(row[0]).strip():
            return

        drum_type = str(row[0]).strip()
        dimensions = str(row[1]).strip() if len(row) > 1 and row[1] else ''
        cbm = row[2] if len(row) > 2 and row[2] else 0
        tare_kg = row[3] if len(row) > 3 and row[3] else 0
        gross_kg = row[4] if len(row) > 4 and row[4] else 0

        if '卡板' in drum_type:
            return

        skip_keywords = ['货柜', '卡板可以放', '集装箱', '1.0*', '1.1*', '20尺']
        if any(kw in drum_type for kw in skip_keywords):
            return

        net_kg_match = re.match(r'(\d+(?:\.\d+)?)\s*kg', drum_type)
        net_kg = float(net_kg_match.group(1)) if net_kg_match else 0

        name = drum_type
        if '细口' in dimensions:
            name = name.replace('50kg蓝桶', '50kg蓝桶(细口)')
        elif '大口' in dimensions:
            name = name.replace('50kg蓝桶', '50kg蓝桶(大口)')

        try:
            cbm_val = float(cbm) if cbm else 0
            tare_val = float(tare_kg) if tare_kg else 0
            gross_val = float(gross_kg) if gross_kg else 0
        except (ValueError, TypeError):
            return

        net = net_kg if net_kg > 0 else (gross_val - tare_val)

        self.packages.append({
            'name': name,
            'dims': dimensions,
            'cbm': cbm_val,
            'tare_kg': tare_val,
            'gross_kg': gross_val,
            'net_kg': net,
        })

    def _load_from_json_file(self, filepath: str):
        """从JSON文件加载包装资料"""
        import json

        # 尝试多种编码
        data = None
        for encoding in ['utf-8', 'gbk', 'gb2312']:
            try:
                with open(filepath, 'r', encoding=encoding) as f:
                    data = json.load(f)
                break
            except UnicodeDecodeError:
                continue

        if data is None:
            raise Exception("无法解码文件")

        # 第一遍：解析桶类型和卡板
        for item in data:
            drum_type = item.get('drum_type', '')
            dimensions = item.get('dimensions', '')
            cbm = item.get('cbm', 0)
            tare_kg = item.get('packing_tare_kg', 0)
            gross_kg = item.get('product_plus_packing_kg', 0)

            # 跳过表头、空行、注释行
            if not drum_type or drum_type in ['种类', '(1) 不打卡板的情况下：', '(2) 打卡板', '(3) 打卡板参考数量']:
                continue

            # 解析卡板
            if '卡板' in drum_type and ('1000' in dimensions or '1100' in dimensions):
                nums = re.findall(r'\d+', dimensions)
                if len(nums) >= 2:
                    size_m = f"{float(nums[0])/1000:.1f}*{float(nums[1])/1000:.1f}"
                    self.pallets.append({
                        'name': f"{size_m}m卡板",
                        'size_m': size_m,
                        'tare_kg': tare_kg,
                        'cbm': 0.15 if '1.0' in size_m else 0.2,
                    })
                continue

            skip_keywords = ['货柜', '卡板可以放', '集装箱', '1.0*', '1.1*', '20尺']
            if any(kw in drum_type for kw in skip_keywords):
                continue

            if not (re.match(r'^\d+kg', drum_type) or drum_type == '1吨桶'):
                continue

            net_kg_match = re.match(r'(\d+(?:\.\d+)?)\s*kg', drum_type)
            net_kg = float(net_kg_match.group(1)) if net_kg_match else 0

            name = drum_type
            if '细口' in dimensions:
                name = name.replace('50kg蓝桶', '50kg蓝桶(细口)')
            elif '大口' in dimensions:
                name = name.replace('50kg蓝桶', '50kg蓝桶(大口)')

            self.packages.append({
                'name': name,
                'dims': dimensions,
                'cbm': cbm,
                'tare_kg': tare_kg,
                'gross_kg': gross_kg,
                'net_kg': net_kg if net_kg > 0 else gross_kg - tare_kg,
            })

        if not self.pallets:
            self.pallets = [
                {'name': '1.0*1.0m卡板', 'size_m': '1.0*1.0', 'tare_kg': 17.0, 'cbm': 0.15},
                {'name': '1.1*1.1m卡板', 'size_m': '1.1*1.1', 'tare_kg': 18.5, 'cbm': 0.2},
            ]

        self._parse_reference_data(data)

        if not self.pallet_capacity:
            self._set_default_pallet_capacity()

    def _parse_reference_data(self, data: list):
        """解析被跳过的参考数据（卡板容量、集装箱尺寸、不打卡板容量）"""
        for item in data:
            drum_type = item.get('drum_type', '')

            # 1. 解析卡板容量数据（打卡板参考数量）
            # 格式: "50kg：1.1*1.1M卡板可以放18桶" 或 "125kg：1.0*1.0M卡板可以放4桶，1.1*1.1M卡板可以放5桶"
            if re.match(r'^\d+kg', drum_type) and '卡板可以放' in drum_type:
                self._parse_pallet_capacity_entry(drum_type)
                continue

            # 2. 解析集装箱内部尺寸
            # 格式: "20gp集装箱内部尺寸:内长5.898米、内宽2.352米、内高2.385米、门高2.28米、门宽2.343米"
            if '集装箱内部尺寸' in drum_type:
                match = re.search(r'内长(\d+\.?\d*)米、内宽(\d+\.?\d*)米、内高(\d+\.?\d*)米', drum_type)
                if match:
                    length = float(match.group(1))
                    width = float(match.group(2))
                    height = float(match.group(3))
                    # 尝试提取门宽门高
                    door_match = re.search(r'门高(\d+\.?\d*)米、门宽(\d+\.?\d*)米', drum_type)
                    door_height = float(door_match.group(1)) if door_match else height
                    door_width = float(door_match.group(2)) if door_match else width

                    self.container_specs['20gp'].update({
                        'length_m': length,
                        'width_m': width,
                        'height_m': height,
                        'door_height_m': door_height,
                        'door_width_m': door_width,
                        'max_cbm': length * width * height,
                    })
                continue

            # 3. 解析不打卡板时货柜容量
            # 格式: "125kg: 20'货柜可以装116桶，14.5吨，每层58桶"
            #       "200kg：20'货柜可以装80桶，16吨，每层40桶"
            #       "1000 IBC吨桶：20'货柜可以装20桶"
            if re.match(r'^\d+kg', drum_type) or 'IBC吨桶' in drum_type:
                if '货柜可以装' in drum_type:
                    match = re.search(r"(\d+)\s*(?:kg|IBC).*?(\d+)\s*桶", drum_type)
                    if match:
                        weight_kg = match.group(1)
                        drums_count = int(match.group(2))
                        # 构建桶类型名（去重：125kg对应125kg新款胶桶）
                        drum_name = self._map_weight_to_drum_name(weight_kg, drum_type)
                        if drum_name:
                            self.no_pallet_container_capacity[drum_name] = (
                                drums_count,
                                f"20'货柜可以装{drums_count}桶"
                            )
                        # 特殊处理1吨桶
                        if 'IBC' in drum_type or weight_kg == '1000':
                            self.no_pallet_container_capacity['1吨桶'] = (
                                drums_count,
                                f"20'货柜可以装{drums_count}桶"
                            )
                    continue

    def _parse_pallet_capacity_entry(self, text: str):
        """解析单条卡板容量数据"""
        # 匹配 "50kg：1.1*1.1M卡板可以放18桶"
        # 也可能 "125kg：1.0*1.0M卡板可以放4桶，1.1*1.1M卡板可以放5桶"
        pattern = r'(\d+)\s*kg[^：]*?：'
        weight_match = re.search(pattern, text)
        if not weight_match:
            return

        weight_kg = weight_match.group(1)
        drum_name = self._map_weight_to_drum_name(weight_kg, text)
        if not drum_name:
            return

        # 如果上下文中包含"纸桶"，也需要为纸桶类型添加条目
        if '纸桶' in text and weight_kg == '50':
            drum_name_paper = '50kg纸桶'
            if drum_name_paper not in self.pallet_capacity:
                self.pallet_capacity[drum_name_paper] = {}
        elif '纸桶' in text and weight_kg == '25':
            drum_name_paper = '25kg 纸桶'
            if drum_name_paper not in self.pallet_capacity:
                self.pallet_capacity[drum_name_paper] = {}

        # 25kg/包（编织袋）和25kg纸桶共享相同的卡板容量
        if weight_kg == '25' and ('纸桶' in text or '编织' in text or '袋' in text):
            # 25kg/包 复用纸桶的容量
            if '25kg/包' not in self.pallet_capacity:
                self.pallet_capacity['25kg/包'] = {}

        # 确保该桶类型在pallet_capacity中有条目
        if drum_name not in self.pallet_capacity:
            self.pallet_capacity[drum_name] = {}

        # 解析所有 "X.X*X.XM卡板可以放Y桶" 或 "X.X*X.Xm卡板可以放Y桶" 格式（大小写均可）
        capacity_pattern = r'(\d\.\d)\*(\d\.\d)[Mm]卡板可以放(\d+)桶'
        for match in re.finditer(capacity_pattern, text):
            size_key = f"{match.group(1)}*{match.group(2)}"
            count = int(match.group(3))
            # 同时更新蓝桶和纸桶（如果上下文中提到纸桶，它们可能共享相同容量）
            self.pallet_capacity[drum_name][size_key] = (
                count,
                f"{size_key}M卡板可以放{count}桶"
            )
            if '纸桶' in text:
                # 纸桶类型
                paper_drum = '50kg纸桶' if weight_kg == '50' else '25kg 纸桶'
                if paper_drum in self.pallet_capacity:
                    self.pallet_capacity[paper_drum][size_key] = (
                        count,
                        f"{size_key}M卡板可以放{count}桶"
                    )
                # 25kg纸桶的容量也要共享给25kg/包（编织袋）
                if weight_kg == '25' and '25kg/包' in self.pallet_capacity:
                    self.pallet_capacity['25kg/包'][size_key] = (
                        count,
                        f"{size_key}M卡板可以放{count}桶"
                    )

    def _map_weight_to_drum_name(self, weight_kg: str, context_text: str = '') -> str:
        """将重量字符串映射到实际的桶类型名称

        Args:
            weight_kg: 重量字符串，如 "30", "50", "125"
            context_text: 上下文文本，用于判断是蓝桶还是纸桶等
        """
        weight = int(weight_kg)
        if weight == 30:
            return '30kg蓝桶'
        elif weight == 50:
            # 需要根据上下文判断是蓝桶还是纸桶
            if '纸桶' in context_text:
                return '50kg纸桶'
            # 默认返回细口蓝桶（50kg蓝桶在JSON中一般指细口）
            # 大口蓝桶也使用相同的卡板容量
            return '50kg蓝桶(细口)'
        elif weight == 25:
            if '纸桶' in context_text:
                return '25kg 纸桶'
            return '25kg/包'
        elif weight == 125:
            return '125kg新款胶桶'
        elif weight == 60:
            return '60kg蓝桶'
        elif weight == 150:
            return '150kg新款胶桶'
        elif weight == 200:
            return '200kg双环闭口桶'
        elif weight == 1000 or '吨桶' in context_text:
            return '1吨桶'
        return None

    def _set_default_pallet_capacity(self):
        """设置默认卡板容量（当JSON解析失败时）"""
        self.pallet_capacity = {
            '30kg蓝桶': {'1.0*1.0': (None, '需实际测量'), '1.1*1.1': (24, '1.1*1.1M卡板可以放24桶')},
            '25kg/包': {'1.0*1.0': (None, '需实际测量'), '1.1*1.1': (18, '1.1*1.1m卡板可以放18桶')},
            '25kg 正方罐（蓝色）': {'1.0*1.0': (None, '需实际测量'), '1.1*1.1': (None, '需实际测量')},
            '25kg 纸桶': {'1.0*1.0': (None, '需实际测量'), '1.1*1.1': (18, '1.1*1.1m卡板可以放18桶')},
            '50kg蓝桶(细口)': {'1.0*1.0': (None, '需实际测量'), '1.1*1.1': (18, '1.1*1.1M卡板可以放18桶')},
            '50kg蓝桶(大口)': {'1.0*1.0': (None, '需实际测量'), '1.1*1.1': (18, '1.1*1.1M卡板可以放18桶')},
            '50kg纸桶': {'1.0*1.0': (None, '需实际测量'), '1.1*1.1': (12, '1.1*1.1m卡板可以放12桶')},
            '60kg蓝桶': {'1.0*1.0': (None, '需实际测量'), '1.1*1.1': (None, '需实际测量')},
            '125kg新款胶桶': {'1.0*1.0': (4, '1.0*1.0M卡板可以放4桶'), '1.1*1.1': (5, '1.1*1.1M卡板可以放5桶')},
            '150kg新款胶桶': {'1.0*1.0': (None, '需实际测量'), '1.1*1.1': (None, '需实际测量')},
            '200kg双环闭口桶': {'1.0*1.0': (None, '需实际测量'), '1.1*1.1': (None, '需实际测量')},
            '1吨桶': {'1.0*1.0': (1, '整箱'), '1.1*1.1': (1, '整箱')},
        }

    def _set_defaults(self):
        """设置默认数据（完整默认值，当JSON文件不存在时）"""
        self.packages = [
            {'name': '30kg蓝桶', 'dims': '250*250*430mm', 'cbm': 0.026, 'tare_kg': 2.0, 'gross_kg': 32.0, 'net_kg': 30.0},
            {'name': '25kg/包', 'dims': '200*500*800mm', 'cbm': 0.028, 'tare_kg': 0.5, 'gross_kg': 25.5, 'net_kg': 25.0},
            {'name': '25kg 正方罐（蓝色）', 'dims': '270*270*410mm', 'cbm': 0.03, 'tare_kg': 2.0, 'gross_kg': 27.0, 'net_kg': 25.0},
            {'name': '25kg 纸桶', 'dims': '310*310*400mm', 'cbm': 0.039, 'tare_kg': 2.0, 'gross_kg': 27.0, 'net_kg': 25.0},
            {'name': '50kg蓝桶(细口)', 'dims': '395*395*585mm', 'cbm': 0.09, 'tare_kg': 2.5, 'gross_kg': 52.5, 'net_kg': 50.0},
            {'name': '50kg蓝桶(大口)', 'dims': '330*390*590mm', 'cbm': 0.08, 'tare_kg': 2.5, 'gross_kg': 52.5, 'net_kg': 50.0},
            {'name': '50kg纸桶', 'dims': '410*410*高500mm', 'cbm': 0.085, 'tare_kg': 2.5, 'gross_kg': 52.5, 'net_kg': 50.0},
            {'name': '60kg蓝桶', 'dims': '320*410*640mm', 'cbm': 0.09, 'tare_kg': 3.5, 'gross_kg': 63.5, 'net_kg': 60.0},
            {'name': '125kg新款胶桶', 'dims': '510*510*810mm', 'cbm': 0.21, 'tare_kg': 6.0, 'gross_kg': 131.0, 'net_kg': 125.0},
            {'name': '150kg新款胶桶', 'dims': '450*450*970mm', 'cbm': 0.196, 'tare_kg': 9.0, 'gross_kg': 159.0, 'net_kg': 150.0},
            {'name': '200kg双环闭口桶', 'dims': '590*590*930mm', 'cbm': 0.31, 'tare_kg': 10.0, 'gross_kg': 210.0, 'net_kg': 200.0},
            {'name': '1吨桶', 'dims': '1200*1000*1150mm', 'cbm': 1.38, 'tare_kg': 58.0, 'gross_kg': 1058.0, 'net_kg': 1000.0},
        ]
        self.pallets = [
            {'name': '1.0*1.0m卡板', 'size_m': '1.0*1.0', 'tare_kg': 17.0, 'cbm': 0.15},
            {'name': '1.1*1.1m卡板', 'size_m': '1.1*1.1', 'tare_kg': 18.5, 'cbm': 0.2},
        ]
        self._set_default_pallet_capacity()
        self.no_pallet_container_capacity = {
            '125kg新款胶桶': (116, "20'货柜可以装116桶"),
            '200kg双环闭口桶': (80, "20'货柜可以装80桶"),
            '1吨桶': (20, "20'货柜可以装20桶"),
        }

    def get_package_options(self) -> list:
        """获取所有包装类型选项"""
        return [p['name'] for p in self.packages]

    def get_pallet_options(self) -> list:
        """获取所有卡板类型选项"""
        return [p['name'] for p in self.pallets]

    def find_package(self, name: str) -> dict:
        """查找包装类型"""
        for pkg in self.packages:
            if name in pkg['name'] or pkg['name'] in name:
                return pkg
        return None

    def find_pallet(self, name: str) -> dict:
        """查找卡板类型"""
        for pallet in self.pallets:
            if name in pallet['name'] or pallet['name'] in name:
                return pallet
        return None

    def calculate(self, drum_name: str, pallet_name: str, total_quantity_kg: float) -> dict:
        """计算包装需求（打卡板模式）"""
        return self.calculate_with_pallet(drum_name, pallet_name, total_quantity_kg)

    def calculate_with_pallet(self, drum_name: str, pallet_name: str, total_quantity_kg: float) -> dict:
        """计算包装需求（打卡板模式）"""
        drum_info = self.find_package(drum_name)
        if not drum_info:
            return {'error': f'未找到桶类型: {drum_name}', 'available': self.get_package_options()}

        pallet_info = self.find_pallet(pallet_name)
        if not pallet_info:
            return {'error': f'未找到卡板类型: {pallet_name}', 'available': self.get_pallet_options()}

        pallet_key = pallet_info['size_m']
        # 查找容量：先精确匹配，如果找不到尝试共享（细口/大口共享容量）
        capacity = self.pallet_capacity.get(drum_info['name'], {}).get(pallet_key)
        if capacity is None:
            # 尝试共享容量：细口和大口视为相同容量
            if '蓝桶' in drum_info['name'] and ('细口' in drum_info['name'] or '大口' in drum_info['name']):
                alt_name = drum_info['name'].replace('细口', '大口').replace('大口', '细口')
                capacity = self.pallet_capacity.get(alt_name, {}).get(pallet_key)
        if capacity is None:
            return {
                'error': f'{drum_info["name"]} 搭配 {pallet_info["name"]}: 需实际测量',
                'drum': drum_info,
                'pallet': pallet_info,
            }
        drums_per_pallet, capacity_note = capacity

        drum_net_kg = drum_info['gross_kg'] - drum_info['tare_kg']
        if drum_net_kg <= 0:
            return {'error': f'桶净重计算错误: {drum_info["name"]}'}

        drums_needed = round(total_quantity_kg / drum_net_kg)
        pallets_needed = (drums_needed + drums_per_pallet - 1) // drums_per_pallet

        total_drums = drums_needed
        total_pallets = pallets_needed

        product_weight = total_quantity_kg
        drum_tare_weight = drum_info['tare_kg'] * total_drums
        pallet_tare_weight = pallet_info['tare_kg'] * total_pallets
        total_tare = drum_tare_weight + pallet_tare_weight
        gross_weight = product_weight + total_tare

        drum_cbm = drum_info['cbm'] * total_drums
        pallet_cbm = pallet_info['cbm'] * total_pallets
        total_cbm = drum_cbm + pallet_cbm

        # 使用动态集装箱参数（20GP）
        container_20gp = self.container_specs['20gp']
        max_cbm = container_20gp['max_cbm']
        max_pallets = container_20gp['max_pallets']

        fits_20gp = total_cbm <= max_cbm and total_pallets <= max_pallets
        full_20gp_loads = max(
            int(total_cbm // max_cbm) + (1 if total_cbm % max_cbm > 0 else 0),
            int(total_pallets // max_pallets) + (1 if total_pallets % max_pallets > 0 else 0)
        )
        full_pallets = total_drums // drums_per_pallet if drums_per_pallet else 0
        remaining_drums = total_drums % drums_per_pallet if drums_per_pallet else 0

        return {
            'input': {
                'drum_type': drum_info['name'],
                'drum_size': drum_info['dims'],
                'drum_net_kg': drum_net_kg,
                'drum_tare_kg': drum_info['tare_kg'],
                'drum_gross_kg': drum_info['gross_kg'],
                'pallet_type': pallet_info['name'],
                'pallet_tare_kg': pallet_info['tare_kg'],
                'total_quantity_kg': total_quantity_kg,
                'capacity_note': capacity_note,
            },
            'result': {
                'total_drums': total_drums,
                'drums_per_pallet': drums_per_pallet,
                'total_pallets': total_pallets,
                'full_pallets': full_pallets,
                'remaining_drums': remaining_drums,
                'product_weight_kg': round(product_weight, 2),
                'drum_tare_kg': round(drum_tare_weight, 2),
                'pallet_tare_kg': round(pallet_tare_weight, 2),
                'total_tare_kg': round(total_tare, 2),
                'gross_weight_kg': round(gross_weight, 2),
                'net_weight_kg': round(product_weight, 2),
                'drum_volume_cbm': round(drum_cbm, 3),
                'pallet_volume_cbm': round(pallet_cbm, 3),
                'total_volume_cbm': round(total_cbm, 3),
            },
            'container_fit': {
                '20gp_limit_cbm': round(max_cbm, 2),
                '20gp_max_pallets': max_pallets,
                'fits_20gp': fits_20gp,
                'full_20gp_loads': full_20gp_loads,
            }
        }

    def calculate_no_pallet(self, drum_name: str, total_quantity_kg: float) -> dict:
        """计算包装需求（不打卡板模式）"""
        drum_info = self.find_package(drum_name)
        if not drum_info:
            return {'error': f'未找到桶类型: {drum_name}', 'available': self.get_package_options()}

        drum_net_kg = drum_info['gross_kg'] - drum_info['tare_kg']
        if drum_net_kg <= 0:
            return {'error': f'桶净重计算错误: {drum_info["name"]}'}

        drums_needed = round(total_quantity_kg / drum_net_kg)
        total_drums = drums_needed

        product_weight = total_quantity_kg
        drum_tare_weight = drum_info['tare_kg'] * total_drums
        total_tare = drum_tare_weight
        gross_weight = product_weight + total_tare

        drum_cbm = drum_info['cbm'] * total_drums
        total_cbm = drum_cbm

        # 使用动态集装箱参数（20GP）
        container_20gp = self.container_specs['20gp']
        max_cbm = container_20gp['max_cbm']

        fits_20gp = total_cbm <= max_cbm
        full_20gp_loads = int(total_cbm // max_cbm) + (1 if total_cbm % max_cbm > 0 else 0)

        return {
            'input': {
                'drum_type': drum_info['name'],
                'drum_size': drum_info['dims'],
                'drum_net_kg': drum_net_kg,
                'drum_tare_kg': drum_info['tare_kg'],
                'drum_gross_kg': drum_info['gross_kg'],
                'total_quantity_kg': total_quantity_kg,
            },
            'result': {
                'total_drums': total_drums,
                'total_pallets': 0,
                'product_weight_kg': round(product_weight, 2),
                'drum_tare_kg': round(drum_tare_weight, 2),
                'total_tare_kg': round(total_tare, 2),
                'gross_weight_kg': round(gross_weight, 2),
                'net_weight_kg': round(product_weight, 2),
                'drum_volume_cbm': round(drum_cbm, 3),
                'pallet_volume_cbm': 0,
                'total_volume_cbm': round(total_cbm, 3),
            },
            'container_fit': {
                '20gp_limit_cbm': round(max_cbm, 2),
                'fits_20gp': fits_20gp,
                'full_20gp_loads': full_20gp_loads,
            }
        }

    def parse_order_requirements(self, order_text: str) -> list:
        """解析订单要求文本"""
        results = []

        # 匹配液体
        liquid_pattern = r'用([一-龥\d]+)\s*(?:共|\s)(\d+)\s*(?:桶|包|袋)'
        matches = re.findall(liquid_pattern, order_text)
        for match in matches:
            package_type = match[0]
            quantity = int(match[1])
            pkg_info = self._detect_package_type(package_type, order_text)
            if pkg_info:
                results.append({
                    'type': '液体',
                    'package': pkg_info['name'],
                    'quantity': quantity,
                    'unit': 'kg',
                    'package_info': pkg_info,
                })

        # 匹配粉
        powder_pattern = r'粉[：:].*?共(\d+)\s*(?:袋|包)'
        powder_match = re.search(powder_pattern, order_text)
        if powder_match:
            powder_qty = int(powder_match.group(1))
            if '编织' in order_text:
                pkg_info = self.find_package('25kg 纸桶')
                if pkg_info:
                    results.append({
                        'type': '粉',
                        'package': pkg_info['name'],
                        'quantity': powder_qty,
                        'unit': 'kg',
                        'package_info': pkg_info,
                    })

        return results

    def _detect_package_type(self, desc: str, full_text: str = '') -> dict:
        """根据描述检测包装类型"""
        if full_text and '大口' in full_text:
            return self.find_package('50kg蓝桶(大口)')

        if '50' in desc and ('蓝桶' in desc or '50kg' in desc or '50公斤' in desc):
            if '大口' in desc:
                return self.find_package('50kg蓝桶(大口)')
            return self.find_package('50kg蓝桶(细口)')

        if '30' in desc and '蓝桶' in desc:
            return self.find_package('30kg蓝桶')

        if '25' in desc:
            if '纸桶' in desc:
                return self.find_package('25kg 纸桶')
            if '正方' in desc or '方罐' in desc:
                return self.find_package('25kg 正方罐（蓝色）')
            if '编织' in desc or '袋' in desc:
                return self.find_package('25kg 纸桶')
            return self.find_package('25kg/包')

        return None