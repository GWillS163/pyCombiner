# ========================  BEGIN inlined module: module1
# ========================  BEGIN inlined module: subdir.deeper
# ========================  BEGIN inlined module: subdir.deeper.deepest
print("This is module4.")
# ========================  END inlined module: subdir.deeper.deepest
print("This is module3.")
# ========================  END inlined module: subdir.deeper
print("This is module1.")
# ========================  END inlined module: module1
# ========================  BEGIN inlined module: subdir
# ========================  BEGIN inlined module: subdir.deeper.module3
# [Already inlined: C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo\subdir\deeper\module3.py]
# ========================  END inlined module: subdir.deeper.module3
print("This is module2.")

# if __name__ == "__main__":
#     print("Other Entrance")
# ========================  END inlined module: subdir
# ========================  BEGIN inlined module: subdir
# [Already inlined: C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo\subdir\deeper\deepest\module4.py]
# ========================  END inlined module: subdir
# ========================  BEGIN inlined module: subdir
# [Already inlined: C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo\subdir\deeper\module3.py]
# ========================  END inlined module: subdir
# ========================  BEGIN inlined module: subdir.deeper.deepest.module4
# [Already inlined: C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo\subdir\deeper\deepest\module4.py]
# ========================  END inlined module: subdir.deeper.deepest.module4
# ========================  BEGIN inlined module: subdir.deeper.deepest
# [Already inlined: C:\Users\PFS\Documents\Codes\pyCombiner\tests\examples\deep_demo\subdir\deeper\deepest\module4.py]
# ========================  END inlined module: subdir.deeper.deepest
from datetime import datetime as t, date
print("This is the main file.")

# if __name__ == "__main__":
#     print("Main Entrance")
