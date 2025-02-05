

10. 局部变量的冲突
5. 处理encoding语句，如果有的话，应该在最前面。


6. 未解析的应该标红
终极目标：
精简版：  导入时，仅导入相关代码

1. Output 美化
```bash
開始ファイルパス: .\batches\inspect_all_table.py
プロジェクトディレクトリ: .
保存ディレクトリ: C:\Users\PFS\PycharmProjects\AutoUITest\batches
File Importing Tree：
└── inspect_all_table.py
    └── operation.py
        └── DriverManager.py
            └── settings.py
                └── resource.py
```

2. 添加更多的setting 控制， 
   3. 比如debug输出
      4. debug内的横线长度不同 表示不同的缩进
   4. 自动排列所有 import 至表头