#!/usr/bin/env python3
"""
使用 Tushare API 获取保利地产(600048.SH)股票数据
"""

import json
import urllib.request
import urllib.error
import time

TUSHARE_TOKEN = "***REMOVED***"
API_URL = "https://api.tushare.pro"

def tushare_query(api_name, params, fields=""):
    """调用 Tushare API"""
    data = {
        "api_name": api_name,
        "token": TUSHARE_TOKEN,
        "params": params,
        "fields": fields
    }
    req = urllib.request.Request(
        API_URL,
        data=json.dumps(data).encode('utf-8'),
        headers={'Content-Type': 'application/json'}
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode('utf-8'))
    except Exception as e:
        return {"code": -1, "msg": str(e)}

def print_section(title):
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")

# 1. 查询股票基本信息
print_section("1. 股票基本信息 (600048.SH 保利地产/保利发展)")
result = tushare_query("stock_basic", {
    "ts_code": "600048.SH",
    "exchange": ""
}, "ts_code,name,area,industry,market,list_date,curr_type")

if result.get("code") == 0:
    data = result.get("data", {})
    fields = data.get("fields", [])
    items = data.get("items", [])
    if items:
        print(f"{' | '.join(fields)}")
        print("-" * 60)
        for item in items:
            print(f"{' | '.join(str(x) for x in item)}")
    else:
        print("未找到数据")
else:
    print(f"错误: {result.get('msg')}")

time.sleep(65)  # 等待避免频率限制

# 2. 查询每日行情数据（最近30天）
print_section("2. 每日行情数据 (最近30天)")
result = tushare_query("daily", {
    "ts_code": "600048.SH",
    "start_date": "20250101",
    "end_date": "20250629"
}, "trade_date,open,high,low,close,vol,amount")

if result.get("code") == 0:
    data = result.get("data", {})
    fields = data.get("fields", [])
    items = data.get("items", [])
    if items:
        print(f"{' | '.join(fields)}")
        print("-" * 60)
        for item in items[:10]:  # 只显示最近10条
            print(f"{' | '.join(str(x) for x in item)}")
        if len(items) > 10:
            print(f"... (共 {len(items)} 条记录，显示前10条)")
    else:
        print("未找到数据")
else:
    print(f"错误: {result.get('msg')}")

time.sleep(65)

# 3. 查询每日指标（市盈率、市净率等）
print_section("3. 每日指标 (daily_basic - 市盈率/市净率等)")
result = tushare_query("daily_basic", {
    "ts_code": "600048.SH",
    "start_date": "20250601",
    "end_date": "20250629"
}, "trade_date,pe_ttm,pb,ps_ttm,dv_ttm,total_mv,circ_mv")

if result.get("code") == 0:
    data = result.get("data", {})
    fields = data.get("fields", [])
    items = data.get("items", [])
    if items:
        print(f"{' | '.join(fields)}")
        print("-" * 60)
        for item in items:
            print(f"{' | '.join(str(x) for x in item)}")
    else:
        print("未找到数据")
else:
    print(f"错误: {result.get('msg')}")

time.sleep(65)

# 4. 查询财务指标（利润表）
print_section("4. 财务指标 (最近4个季度)")
result = tushare_query("income", {
    "ts_code": "600048.SH",
    "start_date": "20240101",
    "end_date": "20251231"
}, "end_date,revenue,total_profit,net_profit,eps")

if result.get("code") == 0:
    data = result.get("data", {})
    fields = data.get("fields", [])
    items = data.get("items", [])
    if items:
        print(f"{' | '.join(fields)}")
        print("-" * 60)
        for item in items[:4]:  # 最近4个季度
            print(f"{' | '.join(str(x) for x in item)}")
    else:
        print("未找到数据")
else:
    print(f"错误: {result.get('msg')}")

print("\n" + "="*60)
print("  数据查询完成")
print("="*60)
