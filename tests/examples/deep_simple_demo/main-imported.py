import module1
from subdir import module2
from subdir import deeper
import subdir.deeper.deepest.module4
import subdir.deeper.deepest
from datetime import datetime as t, date
print("This is the main file.")

if __name__ == "__main__":
    print("Main Entrance")