# 图书电商线上活动价格自动生成器

Windows 桌面版表格处理工具，免费开源，许可证为 MIT。

## 主要功能

- 自动识别 `商品重量`，新增 `快递费` 和 `售价`。
- 支持在界面填写快递费重量区间返回值。
- 支持普通商品和组合商品使用两套售价规则。
- 售价公式：`售价 = (成本价 + 快递费) / (1 - 毛利率)`。
- 价格库默认保存在 `%LOCALAPPDATA%\图书电商线上活动价格自动生成器\price_cache.json`。
- 可将单品表格或组合商品表格计算后的 `商品编码 -> 售价` 保存到价格库。
- 可导入新表格，用 `规格编码` 匹配价格库商品编码，并在表格末尾新增 `售价`。
- 匹配时忽略大小写，并自动截断 `规格编码` 中 `$$` 及其后面的内容。
- 删除行规则由 `rules.json` 控制，配置格式保持兼容。
- 支持 `.xlsx`、`.xlsm`、`.csv`。

## 开发运行

```powershell
py -3.14 main.py
```

## 测试

```powershell
py -3.14 -m pytest
py -3.14 -m py_compile .\main.py
```

## 生成程序

```powershell
.\build_exe.ps1
```

输出：

```text
dist\图书电商线上活动价格自动生成器.exe
```

## 生成 Windows 安装包

```powershell
.\build_installer.ps1
```

输出：

```text
installer\图书电商线上活动价格自动生成器_安装程序.exe
```

## 项目结构

```text
main.py                  启动入口
price_generator\ui.py    Tkinter 界面
price_generator\spreadsheet.py  表格读写与导入导出
price_generator\pricing.py      快递费和售价计算
price_generator\cache.py        价格库读写、导出、清空
price_generator\rules.py        删除规则判断
price_generator\headers.py      表头和别名匹配
price_generator\paths.py        资源、规则、价格库路径
```
