# Phase 1 Tests
# -*- coding: utf-8 -*-
"""
Phase 1 Core Module Tests
Tests: OrderParser, PIExtractor, CodeMatcher, PackageCalculator, Merger
"""
import unittest


class TestOrderParser(unittest.TestCase):
    """OrderParser Tests"""

    def setUp(self):
        from core.order_parser import OrderParser
        self.parser = OrderParser()

    def test_parse_tab_separated(self):
        """Test Tab separated data parsing"""
        test_data = '''王小明\t客户001\t1011\t固色剂\t固色剂\t25kg/桶\t1000\t否\t无\t液体包装+纸桶\t2024-12-15\t已审核\t华南\tHT2024001\t宏昊\tPI\t李娜\t2024-12-01\t确认\t2024-12-20\t出货\t海运\t粉包装异常'''
        result = self.parser.parse(test_data)
        self.assertEqual(len(result), 23)

    def test_validate_valid(self):
        """Test valid data validation"""
        test_data = '''王小明\t客户001\t1011\t固色剂\t固色剂\t25kg/桶\t1000\t否\t无\t液体包装+纸桶\t2024-12-15\t已审核\t华南\tHT2024001\t宏昊\tPI\t李娜\t2024-12-01\t确认\t2024-12-20\t出货\t海运\t粉包装异常'''
        self.parser.parse(test_data)
        is_valid, msg = self.parser.validate()
        self.assertTrue(is_valid)

    def test_get_internal_code(self):
        """Test internal code extraction"""
        test_data = '''王小明\t客户001\t1011\t固色剂\t固色剂\t25kg/桶\t1000\t否\t无\t液体包装+纸桶\t2024-12-15\t已审核\t华南\tHT2024001\t宏昊\tPI\t李娜\t2024-12-01\t确认\t2024-12-20\t出货\t海运\t粉包装异常'''
        self.parser.parse(test_data)
        internal_code = self.parser.get_internal_code()
        self.assertEqual(internal_code, '1011')


class TestCodeMatcher(unittest.TestCase):
    """CodeMatcher Tests"""

    def setUp(self):
        from core.code_matcher import CodeMatcher
        self.matcher = CodeMatcher('knowledge/products_knowledge.json')

    def test_load_knowledge_base(self):
        """Test knowledge base loading"""
        loaded = self.matcher.load('knowledge/products_knowledge.json')
        self.assertTrue(loaded)

    def test_get_version(self):
        """Test version retrieval"""
        self.matcher.load('knowledge/products_knowledge.json')
        version = self.matcher.get_version()
        self.assertIsNotNone(version)

    def test_get_all_info(self):
        """Test product info lookup"""
        self.matcher.load('knowledge/products_knowledge.json')
        info = self.matcher.get_all_info('K70-EB')
        self.assertIsNotNone(info)
        self.assertGreater(len(info), 0)


class TestPackageCalculator(unittest.TestCase):
    """PackageCalculator Tests"""

    def setUp(self):
        from core.package_calculator import PackageCalculator
        self.calc = PackageCalculator('knowledge/packaging_data.json')

    def test_get_package_options(self):
        """Test package options retrieval"""
        options = self.calc.get_package_options()
        self.assertGreater(len(options), 0)
        self.assertIn('30kg蓝桶', options)

    def test_get_pallet_options(self):
        """Test pallet options retrieval"""
        options = self.calc.get_pallet_options()
        self.assertGreater(len(options), 0)

    def test_calculate_no_pallet(self):
        """Test no pallet calculation"""
        result = self.calc.calculate_no_pallet('25kg/包', 500)
        self.assertNotIn('error', result)

    def test_parse_order_requirements(self):
        """Test order requirements parsing"""
        req = '液体包装+纸桶'
        parsed = self.calc.parse_order_requirements(req)
        self.assertIsNotNone(parsed)


class TestMerger(unittest.TestCase):
    """Merger Tests"""

    def test_merge_data(self):
        """Test data merging"""
        from core.merger import Merger
        from core.code_matcher import CodeMatcher

        merger = Merger()

        order_text = '''王小明\t客户001\t1011\t固色剂\t固色剂\t25kg/桶\t1000\t否\t无\t液体包装+纸桶\t2024-12-15\t已审核\t华南\tHT2024001\t宏昊\tPI\t李娜\t2024-12-01\t确认\t2024-12-20\t出货\t海运\t粉包装异常'''
        code_matcher = CodeMatcher('knowledge/products_knowledge.json')
        code_matcher.load('knowledge/products_knowledge.json')

        result = merger.merge(order_text, None, code_matcher)
        self.assertIsNotNone(result)
        self.assertEqual(result.get('内部编号'), '1011')
        self.assertEqual(result.get('产品中文名'), '固色剂')


if __name__ == '__main__':
    unittest.main()