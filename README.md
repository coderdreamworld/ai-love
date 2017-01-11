# xls2lua
一个将excel表格数据导出成lua配置文件的小工具<br>
祖传导表工具是vba写的，vba代码比较难写不想维护了，怒而用py照着写了一个

#特点
视单元格为最小粒度数据（就是说一般不在单元格内定义复合数据），在表头定义数据结构，解析时需要对表头进行递归解析生成类型树<br>
为了让表格的数据描述能力比较强，又能尽量跟excel结合（在单元格中定义复杂数据，会丧失excel的拉数据功能）

# 使用方法
python xls2lua.py < excel文件名> <输出文件名>

# 运行环境
python 2

# 数据格式
## 注释行
* 以双反斜杠「//」开头的行将被忽略
* 空白行会被忽略

## 数据类型
* int 整形<br>
* number 浮点 <br>
* string 字符串<br>
  单元格文本会被添加双引号，且遇到「"」「\n」「\r」会自动加斜杠
* table 单列表<br>
  table类型描述的数据将会在首尾添加一对大括号变成这样：「{<单元格内容>}」
* array 数组 以「array<」开始，「>」结束<br>
* dict 字典 以「dict<」开始，「>」结束<br>
* function 函数<br>
  function(xxx)类型描述的数据会输出成「function(xxx) return <单元格内容> end」，当中function(xxx)是你列头写的
  
## 数据属性
* unique 不能重复
* required 必须有值

# 其他
支持多sheet<br>
