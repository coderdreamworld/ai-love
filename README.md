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
* bool 布尔<br>
  0和false会翻译成false，其余值都会翻译成true
* table 单列表<br>
  table类型描述的数据将会在首尾添加一对大括号变成这样：「{<单元格内容>}」
* array 数组 以「array<」开始，「>」结束<br>
* dict 字典 以「dict<」开始，「>」结束<br>
* function 函数<br>
  function(xxx)类型描述的数据会输出成「function(xxx) return <单元格内容> end」，当中function(xxx)是你列头写的  
  
## 数据修饰器
* unique 不能重复，全局生效
* required 必须有值，只在同一个sheet内生效

## 名字
* 数组容器的元素名字不需要填写名字
* 其余子项必须填写名字
* 路径<br>
  以变量路径为名字的例如「path1.path2.id」，将会从当前解析层级开始依次查找字典容器「path1」、「path2」，在「path2」容器中插入子项「id」

# 其他
支持多sheet<br>
