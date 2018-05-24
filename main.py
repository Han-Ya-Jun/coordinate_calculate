import os
from flask import Flask, render_template, send_from_directory, request, jsonify, make_response
import pandas as pd
from pandas import DataFrame
import time, json, requests

app = Flask(__name__)

UPLOAD_FOLDER = 'upload'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER  # 设置文件上传的目标文件夹
basedir = os.path.abspath(os.path.dirname(__file__))  # 获取当前项目的绝对路径
ALLOWED_EXTENSIONS = set(['txt', 'png', 'jpg', 'xls', 'JPG', 'PNG', 'xlsx', 'gif', 'GIF'])  # 允许上传的文件后缀


# 判断文件是否合法
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS


# 具有上传功能的页面
@app.route('/coordinate')
def upload_test():
    return render_template('upload.html')


@app.route('/sql')
def sql_test():
    return render_template('sql.html')


@app.route('/get_sql',methods=['POST'], strict_slashes=False)
def get_sql():
    file_dir = os.path.join(basedir, app.config['UPLOAD_FOLDER'])  # 拼接成合法文件夹地址
    if not os.path.exists(file_dir):
        os.makedirs(file_dir)  # 文件夹不存在就创建
    f = request.files['myfile']  # 从表单的file字段获取文件，myfile为该表单的name值
    if f and allowed_file(f.filename):  # 判断是否是允许上传的文件类型
        fname = f.filename
        print(fname)
        ext = fname.rsplit('.', 1)[1]  # 获取文件后缀
        table = "`" + str(fname.rsplit('.', 1)[0]) + "`"
        unix_time = int(time.time())
        new_filename = str(unix_time) + '.' + ext  # 修改文件名
        f.save(os.path.join(file_dir, new_filename))  # 保存文件到upload目录
        omcs = pd.read_excel(file_dir + "/" + new_filename)
        omcs_columns = omcs.columns.values.tolist()
        key = ""
        co = 0
        for c in omcs_columns:
            key = key + str(c)
            if co < len(omcs_columns) - 1:
                key = key + ",\n"
            co = co + 1
        omcs_lines = omcs.to_json(orient="records")
        res = json.loads(omcs_lines)
        sql = "-- ----------------------------\n--  insert data to " + table + "\n-- ----------------------------\nBEGIN;\nINSERT INTO " + table + " \n(\n" + key + ")\n VALUES "
        count = 0
        for o in res:
            s = ""
            co = 0
            for c in omcs_columns:
                if o[c] == None:
                    o[c] = " "
                s = s + "'" + str(o[c]) + "'"
                if co < len(omcs_columns) - 1:
                    s = s + ",\n"
                co = co + 1
            sql = sql + "\n(\n" + s + "\n)"
            if count < len(res) - 1:
                sql = sql + ","
            else:
                sql = sql + ";"
            count = count + 1
        sql += "\nCOMMIT;"
        unix_time = int(time.time())
        new_filename = str(unix_time) + '.' + "sql"
        fh = open(file_dir+"/"+new_filename, 'w')
        fh.write(sql)
        fh.close()
        dirpath = os.path.join(app.root_path, 'upload')
        return send_from_directory(dirpath, new_filename, as_attachment=True)  # as_attachment=True 一定要写，不然会变成打开，而不是下载


@app.route('/api/upload', methods=['POST'], strict_slashes=False)
def api_upload():
    file_dir = os.path.join(basedir, app.config['UPLOAD_FOLDER'])  # 拼接成合法文件夹地址
    if not os.path.exists(file_dir):
        os.makedirs(file_dir)  # 文件夹不存在就创建
    f = request.files['myfile']  # 从表单的file字段获取文件，myfile为该表单的name值
    if f and allowed_file(f.filename):  # 判断是否是允许上传的文件类型
        fname = f.filename
        print(fname)
        ext = fname.rsplit('.', 1)[1]  # 获取文件后缀
        unix_time = int(time.time())
        new_filename = str(unix_time) + '.' + ext  # 修改文件名
        f.save(os.path.join(file_dir, new_filename))  # 保存文件到upload目录
        d = pd.read_excel(file_dir + "/" + new_filename)
        addr = d.to_json(orient='records')
        address_json = json.loads(addr)
        addresses = []
        for ad in address_json:
            address = ad["地址"]
            ADDRESS_URL = ""
            url = ADDRESS_URL + str(address)
            req = requests.get(url)
            req = req.json()
            if req["status"] == 0:
                l = [ad["地址"], req["result"]["xcoord"], req["result"]["ycoord"]]
            else:
                l = [ad["地址"], "", ""]
            addresses.append(l)
        df_add = DataFrame(addresses, columns=["地址", "经度", "纬度"])
        unix_time = int(time.time())
        new_filename = str(unix_time) + '.' + "xlsx"
        writer = pd.ExcelWriter(file_dir + "/" + new_filename)
        df_add.to_excel(writer, 'Sheet1')
        writer.save()
        dirpath = os.path.join(app.root_path, 'upload')  # 这里是下在目录，从工程的根目录写起，比如你要下载static/js里面的js文件，这里就要写“static/js”
        return send_from_directory(dirpath, new_filename, as_attachment=True)  # as_attachment=True 一定要写，不然会变成打开，而不是下载
    else:
        return jsonify({"errno": 1001, "errmsg": "上传失败"})


if __name__ == '__main__':
    app.run(host="0.0.0.0", port=8888)
