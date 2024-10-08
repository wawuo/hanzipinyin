# I 汉字拼音表

根据网上资料整理的汉字拼音表，主要包括Unicode 0x4E00—0x9FA5中的汉字，其中一些韩国汉字、日本汉字不包含在内。

## 数据目录

### 原始文本数据

- **hzpy-utf8.txt**：汉字列表，每行6列，分别是：
  1. 汉字本身
  2. 汉字的拼音
  3. 声母
  4. 韵母
  5. Unicode编码
  6. 0表示不常用汉字，1表示常用汉字，2表示该汉字是姓氏

  多音字每个读音单独一行。

- **simplified2traditional.txt**：简繁转换表，第一列是简体Unicode编码，第二列是对应的繁体字Unicode编码。

## 数据库目录

- **hanzi.db**：sqlite3数据库，其中数据位于`hanzi`表中。这个数据库是用`script`中的`store-hanzi.py`创建的，该表共5列，结构如下：
  1. `unicode`：int类型，主键，内容为汉字的Unicode编码
  2. `pinyin`：text类型，汉字的拼音，多音字的多个读音用","(英文逗号)隔开
  3. `type`：int类型，1为简体字，2为繁体字，其他汉字为0
  4. `map`：int类型，简体字对应的繁体字，或繁体字对应的简体字，或0
  5. `freq`：int类型，和hzpy-utf8.txt中的最后一列含义相同，多音字取最大值

## 脚本目录

- **script**：包含一些没什么用的脚本。

# II 使用VBA与Access数表转换汉字拼音的使用说明

## 20240829 使用VBA与Access数表（hanzi）在Excel中使用函数来转换拼音

VBA使用ADO（ActiveX Data Objects）

1. 在Alt+F11中插入一个新模块
2. 在Excel中空白单元格里使用 `=ConvertHanziToPinyin(A1)`，其中A1是汉字单元格坐标

