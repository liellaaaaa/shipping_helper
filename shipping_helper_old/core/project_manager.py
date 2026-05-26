# -*- coding: utf-8 -*-
"""
项目文件夹管理模块
以PI号为项目文件夹名，自动创建Phase子目录结构
"""

import os
import json
import shutil


class ProjectManager:
    """项目文件夹管理器"""

    PHASES = ['p1', 'p2', 'p3', 'p4']

    def __init__(self, base_dir: str = None):
        """
        初始化项目管理器
        :param base_dir: 基础目录，默认为当前目录
        """
        self.base_dir = base_dir or os.getcwd()
        self.project_dir = None

    def create_project(self, pi_no: str) -> str:
        """
        创建项目文件夹
        :param pi_no: PI号（作为项目文件夹名）
        :return: 项目文件夹路径
        """
        # 清理PI号中的非法字符
        pi_no = self._sanitize_filename(pi_no)

        # 创建项目根目录
        self.project_dir = os.path.join(self.base_dir, pi_no)
        os.makedirs(self.project_dir, exist_ok=True)

        # 创建Phase子目录
        for phase in self.PHASES:
            phase_dir = os.path.join(self.project_dir, phase)
            os.makedirs(phase_dir, exist_ok=True)

            # 创建input和output子目录
            input_dir = os.path.join(phase_dir, 'input')
            output_dir = os.path.join(phase_dir, 'output')
            os.makedirs(input_dir, exist_ok=True)
            os.makedirs(output_dir, exist_ok=True)

        return self.project_dir

    def get_project_dir(self, pi_no: str = None) -> str:
        """
        获取项目文件夹路径
        :param pi_no: PI号（可选，不提供则返回当前项目）
        :return: 项目文件夹路径
        """
        if pi_no:
            pi_no = self._sanitize_filename(pi_no)
            return os.path.join(self.base_dir, pi_no)
        return self.project_dir

    def get_phase_dir(self, phase: str, subdir: str = 'output') -> str:
        """
        获取Phase子目录
        :param phase: Phase名称（如 'p1'）
        :param subdir: 子目录名称（'input' 或 'output'）
        :return: 目录路径
        """
        if not self.project_dir:
            return None
        return os.path.join(self.project_dir, phase, subdir)

    def save_json(self, phase: str, filename: str, data: dict) -> str:
        """
        保存JSON文件到Phase的output目录
        :param phase: Phase名称
        :param filename: 文件名
        :param data: 要保存的数据
        :return: 保存的文件路径
        """
        output_dir = self.get_phase_dir(phase, 'output')
        if not output_dir:
            return None

        filepath = os.path.join(output_dir, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        return filepath

    def copy_input_file(self, phase: str, src_filepath: str) -> str:
        """
        复制输入文件到Phase的input目录
        :param phase: Phase名称
        :param src_filepath: 源文件路径
        :return: 目标文件路径
        """
        input_dir = self.get_phase_dir(phase, 'input')
        if not input_dir:
            return None

        filename = os.path.basename(src_filepath)
        dst_filepath = os.path.join(input_dir, filename)

        # 如果文件已存在，先删除
        if os.path.exists(dst_filepath):
            os.remove(dst_filepath)

        shutil.copy2(src_filepath, dst_filepath)
        return dst_filepath

    def project_exists(self, pi_no: str) -> bool:
        """
        检查项目是否存在
        :param pi_no: PI号
        :return: 是否存在
        """
        pi_no = self._sanitize_filename(pi_no)
        project_path = os.path.join(self.base_dir, pi_no)
        return os.path.exists(project_path)

    def _sanitize_filename(self, filename: str) -> str:
        """
        清理文件名中的非法字符
        """
        # Windows非法字符
        illegal_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        for char in illegal_chars:
            filename = filename.replace(char, '_')
        return filename.strip()
