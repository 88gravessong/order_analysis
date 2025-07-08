#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py
简单的 Flask Web 应用：
1. 首页提供多文件上传和日期范围选择（精确到日）。
2. POST /process 接收文件与日期范围，调用 compute_logic.compute_metrics。
3. 返回生成的 Excel 文件供下载。
"""

from datetime import datetime, date
from io import BytesIO
from zipfile import ZipFile
import os
import uuid
from tempfile import gettempdir

from flask import Flask, render_template, request, send_file, flash, redirect, url_for, after_this_request

from compute_logic import compute_metrics

app = Flask(__name__)
app.secret_key = "secret-key-change-me"


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
    # 在本地测试使用，部署时请使用 WSGI Server
    app.run(host="0.0.0.0", port=5000, debug=True) 