#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app_standalone.py
订单分析系统独立版本 - 适用于打包成exe
"""

import os
import sys
import webbrowser
import threading
import time
from datetime import datetime, date
from io import BytesIO
from zipfile import ZipFile
import uuid
from tempfile import gettempdir

from flask import Flask, render_template, request, send_file, flash, redirect, url_for, after_this_request

# 确保在打包后能找到模板文件
if getattr(sys, 'frozen', False):
    # 如果是打包后的exe
    template_dir = os.path.join(sys._MEIPASS, 'templates')
else:
    # 如果是开发环境
    template_dir = 'templates'

app = Flask(__name__, template_folder=template_dir)
app.secret_key = "order-analysis-secret-key-2024"

# 导入计算逻辑
try:
    from compute_logic import compute_metrics
except ImportError:
    # 如果导入失败，定义一个简化版本
    def compute_metrics(file_streams, start_date, end_date):
        from collections import defaultdict
        from openpyxl import Workbook
        return Workbook(), defaultdict(dict)


def open_browser():
    """延迟打开浏览器"""
    time.sleep(1.5)  # 等待Flask启动
    webbrowser.open('http://localhost:5000')


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    if "files" not in request.files:
        flash("请至少上传一个文件！")
        return redirect(url_for("index"))

    files = request.files.getlist("files")
    if not files or files[0].filename == "":
        flash("未选择文件！")
        return redirect(url_for("index"))

    # 日期解析
    try:
        start_str = request.form.get("start_date")
        end_str = request.form.get("end_date")
        start_date = datetime.strptime(start_str, "%Y-%m-%d").date() if start_str else date.min
        end_date = datetime.strptime(end_str, "%Y-%m-%d").date() if end_str else date.max
    except ValueError:
        flash("日期格式错误，应为 YYYY-MM-DD")
        return redirect(url_for("index"))

    if start_date > end_date:
        flash("开始日期不能晚于结束日期！")
        return redirect(url_for("index"))

    # 读取文件到 BytesIO 列表
    file_streams = []
    for f in files:
        data = BytesIO(f.read())
        file_streams.append(data)

    try:
        wb, stats = compute_metrics(file_streams, start_date, end_date)
        if not stats:
            flash("在所选日期范围内未找到符合条件的数据，请调整日期或检查文件！")
            return redirect(url_for("index"))
    except Exception as e:
        flash(f"处理文件时发生错误: {e}")
        return redirect(url_for("index"))

    # 将结果保存到临时文件
    temp_filename = f"order_metrics_{uuid.uuid4().hex[:8]}.xlsx"
    temp_path = os.path.join(gettempdir(), temp_filename)
    wb.save(temp_path)
    
    # 准备结果数据用于前端显示
    results_data = []
    for sku, metrics in sorted(stats.items(), key=lambda x: (-x[1]["total"], x[0])):
        total = metrics["total"]
        if total == 0:
            continue
        completed_rate = metrics.get("completed", 0) / total * 100
        delivered_rate = metrics.get("delivered", 0) / total * 100
        refund_rate = metrics.get("refund", 0) / total * 100
        cancel_before_rate = metrics.get("cancel_before", 0) / total * 100
        cancel_after_rate = metrics.get("cancel_after", 0) / total * 100
        in_transit_rate = metrics.get("in_transit", 0) / total * 100
        sign_rate = completed_rate + delivered_rate + refund_rate
        
        results_data.append({
            'seller_sku': sku,
            'total': total,
            'sign_rate': round(sign_rate, 2),
            'completed_rate': round(completed_rate, 2),
            'delivered_rate': round(delivered_rate, 2),
            'refund_rate': round(refund_rate, 2),
            'cancel_before_rate': round(cancel_before_rate, 2),
            'cancel_after_rate': round(cancel_after_rate, 2),
            'in_transit_rate': round(in_transit_rate, 2),
        })
    
    return render_template("results.html", 
                         results=results_data,
                         start_date=start_date,
                         end_date=end_date,
                         temp_filename=temp_filename,
                         total_files=len(files),
                         total_orders=sum(r['total'] for r in results_data))


@app.route("/download/<filename>")
def download(filename):
    """下载临时生成的结果文件"""
    temp_path = os.path.join(gettempdir(), filename)
    if not os.path.exists(temp_path):
        flash("文件不存在或已过期！")
        return redirect(url_for("index"))
    
    def remove_file():
        """下载后删除临时文件"""
        try:
            os.remove(temp_path)
        except:
            pass
    
    # 使用 Flask 的 after_this_request 在响应后删除文件
    @after_this_request
    def cleanup(response):
        remove_file()
        return response
    
    return send_file(temp_path, 
                    as_attachment=True, 
                    download_name=f"订单指标分析结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    print("=" * 50)
    print("🚀 订单分析系统正在启动...")
    print("📊 Web界面地址: http://localhost:5000")
    print("💡 程序会自动打开浏览器")
    print("⚡ 按 Ctrl+C 停止服务")
    print("=" * 50)
    
    # 在后台线程中打开浏览器
    threading.Thread(target=open_browser, daemon=True).start()
    
    try:
        # 启动Flask应用
        app.run(host="0.0.0.0", port=5000, debug=False)
    except KeyboardInterrupt:
        print("\n👋 感谢使用订单分析系统！")
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        input("按回车键退出...") 