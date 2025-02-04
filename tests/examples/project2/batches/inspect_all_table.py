import os
from datetime import datetime
import psycopg2
from psycopg2 import sql
from tests.util.operation import write_to_result, Color


class DatabaseSearcher:
    def __init__(self, host, dbname, user, password, key_name, values):
        self.host = host
        self.dbname = dbname
        self.user = user
        self.password = password
        self.key_name = key_name
        self.values = values
        self.conn = None
        self.cur = None

        # 设置输出文件路径
        current_date = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        self.output_dir = f"Output/DB [{self.key_name}]_{current_date}-[" + self.values.replace('%', '').replace('\t','')[:10] + "]/"
        os.makedirs(self.output_dir, exist_ok=True)  # 如果文件夹不存在，创建文件夹

    def connect(self):
        """连接数据库"""
        self.conn = psycopg2.connect(
            host=self.host, dbname=self.dbname, user=self.user, password=self.password
        )
        self.cur = self.conn.cursor()

    def disconnect(self):
        """关闭数据库连接"""
        if self.cur:
            self.cur.close()
        if self.conn:
            self.conn.close()

    def output_file_print(self, results, table_name):
        """打印和写入查询结果"""
        file_path = os.path.join(self.output_dir, f"{table_name}_({len(results)} items).txt")

        print(f"\t{Color.GREEN}Found matching rows in {table_name}{Color.RESET}:")

        for i, row in enumerate(results):
            if not i > 10:
                print(row)

            write_to_result(str(row), file_path)

        print(f"------Total ({len(results)}) items\n")

    def search(self):
        """搜索数据库中的表和记录"""
        try:
            # 获取所有表的名字
            self.cur.execute("""
                SELECT table_name 
                FROM information_schema.tables 
                WHERE table_schema = 'admin'
            """)
            tables = self.cur.fetchall()

            for table in tables:
                table_name = table[0]
                if "view" in str(table_name):
                    continue

                print(f"Checking table: {table_name}", end=" ")

                # 检查当前表是否有指定的列
                self.cur.execute("""
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_name = %s AND column_name = %s
                """, (table_name, self.key_name))

                columns = self.cur.fetchall()

                if columns:
                    print(
                        f"\nFound '{self.key_name}' in table: {Color.MAGENTA}{table_name: ^20}{Color.RESET}. Querying for {self.values}...",
                        end=" "
                    )
                    if "%" in self.values:
                        query = sql.SQL("SELECT * FROM {} WHERE {} LIKE %s").format(
                            sql.Identifier(table_name),  # 动态插入表名
                            sql.Identifier(self.key_name)  # 动态插入列名
                        )
                    else:
                        query = sql.SQL("SELECT * FROM {} WHERE {} = %s").format(
                            sql.Identifier(table_name),  # 动态插入表名
                            sql.Identifier(self.key_name)  # 动态插入列名
                        )
                    self.cur.execute(query, (self.values,))
                    results = self.cur.fetchall()

                    if results:
                        self.output_file_print(results, table_name)
                    else:
                        print(f"\t{Color.RED}[No matching rows found]{Color.RESET}\n")
                else:
                    print(f"\t\t{Color.RED}[No '{self.key_name}' column]{Color.RESET}\n")

        finally:
            self.disconnect()


# 示例调用
if __name__ == "__main__":
    host = 'localhost'
    dbname = 'bks'
    user = 'admin'
    password = 'admin'
    key_name = 'kanribo_cd_harai'
    values = '1952848' # if "%" in

    searcher = DatabaseSearcher(host, dbname, user, password, key_name, values)
    searcher.connect()
    searcher.search()
