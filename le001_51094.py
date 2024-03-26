import pandas as pd
import pyodbc


def run():

    excel = pd.read_excel('le001_51094.xlsx')

    for row in excel.iterrows():
        # print(row)
        p1 = f"{row[1]['code']:6f}".replace('.','').zfill(8)
        code = f"N{p1}4"
        # print(c)

        # break

        # 定義連線資訊
        server = 'db1'  # 例如 'localhost\sqlexpress'
        database = 'wpap1'
        username = 'iemis'
        password = 'ooooo'

        # 建立連線字符串
        cnxn_str = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        # 或者使用Windows身份驗證的連線字符串
        # cnxn_str = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes'

        # 建立連線
        cnxn = pyodbc.connect(cnxn_str)

        # 創建游標
        cursor = cnxn.cursor()


        # 準備SQL查詢
        sql = """
        INSERT INTO CottonImportTemp_LE001_20240325 (
            Factory, 
            SupplierName, 
            CottonType, 
            CottonSpec, 
            CottonBatchNo, 
            CottonNetWeight, 
            CottonGrossWeight, 
            Weight, 
            IsImported, 
            IsDivided, 
            IsMaterialOut, 
            IsMaterialBack, 
            ImportTime, 
            MoisturRate, 
            PONo, 
            PackingDate, 
            el_no, 
            AcceptDate, 
            su_bno, 
            BarcodeID ) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """

        # 執行SQL查詢
        cursor.execute(sql, 
                    'SP4',
                    'LE001',
                    'R',
                    '1.5/40',
                    code,
                    row[1]['net'],
                    row[1]['gross'],
                    row[1]['cond'],
                    'N',
                    'N',
                    'N',
                    'N',
                    '2024/03/25 11:50:00',
                    '13.0000',
                    'EC24300107',
                    '2024/03/07',
                    '1SRS15D40NP',
                    '2024/03/07',
                    code,
                    code)

        # 提交事務
        cnxn.commit()

        # 關閉游標和連線
        cursor.close()
        cnxn.close()

        print(f"{row[1]['sn']}資料插入成功")

        # break