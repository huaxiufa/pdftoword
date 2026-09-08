# Excel / CSV 清洗器

一个直接可用的本地 Web 小工具：上传 `.xlsx`、`.xlsm` 或 `.csv`，自动清理常见脏数据并下载干净的 `.xlsx`。

## 功能

- 删除整行空白
- 删除重复记录
- 智能按邮箱 / 手机号去重
- 清理前后空格、重复空白、不可见字符、全角字符
- 将 `N/A`、`NULL`、`TBD`、`未知`、`-` 等常见占位值统一为空
- 自动规范邮箱为小写
- 自动规范手机号，只保留数字或国际 `+` 前缀
- 检查明显异常邮箱
- 检查日期列中的明显异常值
- 自动生成 `Cleaned`、`Issues`、`Summary` 三个工作表
- CSV 自动尝试 UTF-8、GB18030、Big5、CP1252 等常见编码
- 原始文件不会被修改

这些清洗项与 Excel 常见的数据清理工作一致：Microsoft 的数据清理建议包括删除重复记录、处理前后空格/不可打印字符以及统一文本大小写。citeturn0search0turn0search1

## 运行

```bash
docker compose up -d --build
```

打开：

```text
http://localhost:8000
```

## 更新

```bash
git pull origin main
docker compose up -d --build
```

## 输出格式

上传 Excel 或 CSV 后，程序始终输出一个 `.xlsx` 文件：

- **Cleaned**：清洗后的数据
- **Issues**：发现并处理/标记的问题明细
- **Summary**：原始行数、清洗后行数、删除重复数、问题数量等统计

CSV 默认读取第一个数据表；Excel 默认处理第一个工作表。

## 商业化建议

适合进一步增加免费额度、付费额度和订阅，例如免费 500 行、单次大文件付费、或按月订阅。Stripe Payment Links 支持一次性付款和订阅，可以作为后续收费入口。citeturn0search0
