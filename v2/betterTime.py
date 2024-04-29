import time


class BetterTime:
    @staticmethod
    def timeName() -> str:
        """
        格式化时间戳为字符串 "YYYYMMDDHHMMSS" 作为文件名

        param:
            None

        return:
            str
        """
        return str(time.strftime("%Y%m%d%H%M%S", time.localtime(time.time())))

    @staticmethod
    def timeLog() -> str:
        """
        格式化时间戳为字符串 "YYYY-MM-DD HH:MM:SS" 作为日志输出时间

        param:
            None

        return:
            str
        """
        return str(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time())))

    @staticmethod
    def timeStamp() -> int:
        """
        获取当前时间戳

        param:
            None

        return:
            int
        """
        return int(time.time())
