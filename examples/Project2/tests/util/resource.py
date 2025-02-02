class Accounts:
    credentials = {
        "本庁": ["10000", "1000010000"],
        "艦艇装備研究所": ["16000", "1600016000"],
        "次世代装備研究所": ["19000", "1900019000"]
    }

    @classmethod
    def get_credentials(cls, name) -> list:
        pair = cls.credentials.get(name)
        if pair is None:
            raise ValueError(f"[ {name} ] not existing")
        return pair # usr, pwd

class Kanri_irai:
    # 【依頼】
    供用_0 = 0
    返納_1 = 1

    # 【受入依頼】
    寄付_2 = 2
    借受_3 = 3
    編入_4 = 4
    転用_5 = 5
    雑件_6 = 6

    # 【生産報告】
    生産_7 = 7
    副生_8 = 8
    発生材_9 = 9
    物品減耗_10 = 10

    # 【不用決定依頼】
    転用_不用_11 = 11
    売払_12 = 12
    解体_13 = 13
    廃棄_14 = 14

    # 【払出依頼】
    交換_15 = 15
    編入_払出_16 = 16
    譲与_17 = 17
    貸付_18 = 18
    亡失_19 = 19
    雑件_払出_20 = 20

    # 【寄託依頼】
    寄託_21 = 21
    物品保管_22 = 22

    # 【管理換】
    管理換_払_23 = 23
    管理換_受_24 = 24
    分類換_払_25 = 25
    分類換_受_26 = 26
    供用換_払_27 = 27
    管理換協議依頼_払_28 = 28
    管理換協議依頼_受_29 = 29
    管理換協議通知_払_30 = 30
    管理換協議通知_受_31 = 31

    # 【依頼差戻し】
    依頼差戻し_32 = 32


class Kyouyou_irai:
    # 【管理官取得】
    新規取得_0 = 0
    返品_材料使用_1 = 1
    交換_2 = 2
    管理換_受_3 = 3
    分類換_受_4 = 4
    管理官取得_雑件_5 = 5

    # 【受入】
    # 【供用受入】
    供用_6 = 6
    供用換_7 = 7
    # 【生産受入】
    生産_8 = 8
    副生_9 = 9
    発生材_10 = 10
    # 【受入】
    寄付_11 = 11
    借受_12 = 12
    編入_13 = 13
    転用_14 = 14
    受入_雑件_15 = 15
