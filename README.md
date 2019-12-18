## data-entry

data-entry是一个可用于从Doc文档、Excel文档等常见的文档中提取结构化信息的工具

## 开发状态

data-entry已经为数据提取工作提供了非常大的便利，实现逻辑简单并易于配置，非常适合做二次开发满足自身数据提取的需求，目前已经根据业务需求做了多次调整，应该能够满足大部分的使用场景。

### 使用示例

```
java -jar data-entry.jar excel docs...
使用系统变量: -DprintDoc=true, 打印实际读取的值
使用系统变量如: -DgetValues=水库简介, 获取水库简介后面的值
列标识: key_index#valueIndex|explain
列标识: key 列头的标识，如：水库名称
列标识: index　当存在多个相同的标识的时候需要使用下标来区分，默认从0开始
列标识: valueIndex 值对应key的位置，默认是行对应(valueIndex从0开始)，也可以从列往下对于（valueIndex从-1开始）
列标识: explain 注释
头必须存在两列，第一列是key的headers，第二行为用户自己指定，但必须存在
headers中添加了两个预定义header，可以直接使用(文件目录、文件名字)
```

- excel为需要提取的数据结构描述文件

- docs为需要提取的文件夹，可以指定多个文件夹

- 结构化数据提取之后保存在excel另外的sheet中

- 例子：

  ```
  java -jar -DprintDoc=true data-entry.jar test.xlsx test.doc
  ```

## 开发计划

计划在后期版本中加入功能和优化

- 支持更加复杂的提取需求
- 支持更多文件类型提取