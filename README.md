# 图书电商线上活动价格自动生成器

Windows 桌面版表格处理工具。

## License

MIT License. Free and open source.

功能：

- 自动识别“商品重量”列
- 在商品重量列旁新增“快递费”列
- 可在界面中填写每个商品重量区间返回的快递费数字
- 可在界面中填写成本价分界值，以及高低两档毛利率
- 删除行规则可通过 `rules.json` 配置，不需要重新打包程序
- 如果同一行“商品名”和“成本价”都为空白，导出时删除该行
- 支持 `.xlsx`、`.xlsm`、`.csv`

生成程序：

```powershell
.\build_exe.ps1
```

生成后程序在：

```text
dist\图书电商线上活动价格自动生成器.exe
```

生成安装程序：

```powershell
.\build_installer.ps1
```

生成后安装包在：

```text
installer\图书电商线上活动价格自动生成器_安装程序.exe
```

删除规则配置：

```text
rules.json
```

程序会默认读取同目录下的 `rules.json`。也可以在主界面选择其它规则文件，或点击“生成”导出一份默认规则。

支持的规则操作：

- `blank`：为空
- `not_blank`：不为空
- `zero_or_blank`：为 0 或为空
- `contains`：包含指定文字
- `contains_ci`：包含指定文字，忽略大小写
- `no_english_letter`：不包含英文字母
- `four_digit_lt`：找到连续 4 位数字且小于指定值
- `lt`、`lte`、`gt`、`gte`、`eq`：数字比较
