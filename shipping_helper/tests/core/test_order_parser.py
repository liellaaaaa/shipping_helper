# shipping_helper/tests/core/test_order_parser.py
import pytest
from core.order_parser import OrderParser


def test_parse_tab_separated():
    parser = OrderParser()
    text = "张三\tWA524\tK70-EB\t皂洗剂\t\t50kg\t1000\t否\t无\t液体：50kg蓝桶细口，共10桶\n粉：25kg纸桶，共5袋\t2026-06-01\t已审核\t华南\tHT260301\t宏昊\tPI\t李四\t2026-03-01\t是\t2026-06-15\t快递\t件\t正常"
    result = parser.parse(text)
    assert result['业务员'] == '张三'
    assert result['客户编号'] == 'WA524'
    assert result['内部编号'] == 'K70-EB'


def test_fix_column_offset():
    parser = OrderParser()
    # 报关名称位置是纯数字的情况
    text = "张三\tWA524\tK70-EB\t皂洗剂\t1000\t50kg\t1000\t否\t无\t备注\t2026-06-01"
    result = parser.parse(text)
    assert result['报关名称'] == ''


def test_empty_text():
    parser = OrderParser()
    result = parser.parse("")
    assert result == {}


def test_validate_empty_pi():
    parser = OrderParser()
    parser.parse("张三\tWA524\t\t皂洗剂")
    is_valid, msg = parser.validate()
    assert is_valid == False
    assert "订单号" in msg