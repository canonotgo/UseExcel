This is a library that stores VBA code.

```excel
SUBSTITUTE:
    function is used to replace occurrences of a specified substring within a text string. The syntax for the SUBSTITUTE function is:
    =SUBSTITUTE(text, old_text, new_text, [instance_num])

LEFT:
    function is used to extract a specified number of characters from the beginning (left side) of a text string. Here's a brief overview of how to use the LEFT function:
    LEFT(text, [num_chars])

Example: Get prefix of mac address
    =LEFT(SUBSTITUTE(B2,":",""),6)

Example: Whether the value is the same
    =IF(D2=F2,"Same","Different")
```

### 统计

```bash
# 通过 INDIRECT 间接的关联B2的值 统计Sheet名称等于 B2的 A列有多少有效数据
=COUNTA(INDIRECT("'"&B2&"'!A:A"))

# 根据 Sheet0 表单的 B2:B1000 范围 统计有数据总和
=COUNTA('Sheet0'!B2:B1000)

# 在 Sheet0 表单 G列 统计值为B3的数据

# 在 Sheet1 表单 A 列 统计值 "Same" 出现的次数。
=COUNTIF(Sheet1!A:A, "Same")

# 统计 Sheet1 表单 A 列 为 "Same" 且 B 列 为 "Type1" 的行数
=COUNTIFS(Sheet1!A:A, "Same", Sheet1!B:B, "Type1")
```

### 查找

```bash
# 在 'C:\Users\cc\Desktop\Excel\[all_data.xlsx]Sheet1'!$A$1:$D$100000 区域 查找值等于 A2 数据，将制定区域的第二列数据 通过 FALSE(精准匹配) 输出
=VLOOKUP(A2,'C:\Users\cc\Desktop\Excel\[all_data.xlsx]Sheet1'!$A$1:$D$100000,2,FALSE)

# 关键字查找 1: 在 Sheet1 表单 创建A(key) B(The value we need), 在当前 Sheet 页想找出包含 Sheet1 A列的关键字, 包含是获取 Sheet1 B列的值
=IF(SUMPRODUCT(--ISNUMBER(SEARCH(Sheet1!$A$2:$A$96, B83))) > 0, INDEX(Sheet1!$B$2:$B$96, MATCH(TRUE, ISNUMBER(SEARCH(Sheet1!$A$2:$A$96, B83)), 0)), "")
```