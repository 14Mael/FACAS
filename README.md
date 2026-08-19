# FACAS 初版

连接已经登录的 Microsoft Edge，读取“销售管理”中的现场销售、库房销售、拆解料销售，按日期范围生成 Excel。

## 使用

1. 使用远程调试端口启动 Edge，并在新窗口登录系统：

   ```powershell
   & "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" `
     --remote-debugging-port=9222 `
     --user-data-dir="$env:TEMP\facas-edge"
   ```

2. 安装依赖：`python -m pip install -r requirements.txt`
3. 运行：`python app.py`
4. 选择日期范围和销售分类，点击“开始提取”。推荐使用 `.un-facas.ps1` 启动，这样不会在工作区生成 Python 缓存文件；命令行参数可直接透传。

每张销售单输出一行，只读取详情中的第一条物料和第一条回收单号；总金额、总重量和备注取销售单主表。
