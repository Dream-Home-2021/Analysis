import base64
import pickle


class AES:
    def __init__(self):
        self.cache_file_dasou = 'cache_0.pkl'
        self.cache_file_stream = 'cache_1.pkl'
        self.cache_file_dasou_stream = 'cache_2.pkl'

    # 加密函数
    def encrypt_data(self, data):
        # 将字典转换为二进制
        binary_data = pickle.dumps(data)
        # 使用base64进行加密
        return base64.b64encode(binary_data)

    # 解密函数
    def decrypt_data(self, encrypted_data):
        # 使用base64进行解密
        decrypted_data = base64.b64decode(encrypted_data)
        # 将二进制转换回字典
        return pickle.loads(decrypted_data)

    # 保存数据到缓存文件
    def save_data_to_cache(self, data, file_name):
        # 加密数据
        encrypted_data = self.encrypt_data(data)
        # 保存到缓存文件
        with open(f'./Cache/{file_name}', 'wb') as f:
            f.write(encrypted_data)

    # 从缓存文件读取数据
    def load_data_from_cache(self, file_name):
        # 读取缓存文件
        with open(f'./Cache/{file_name}', 'rb') as f:
            encrypted_data = f.read()
        # 解密数据
        return self.decrypt_data(encrypted_data)


class StatisticsCalculator(AES):
    def __init__(self, df):
        self.df = df
        self.departments = ["新开部门", "维护部门", "大客部门", "泉州部门", "运营策略中心", "失效挽救部", "品牌部",
                            "框架",
                            "漳州客服部", "行发维护大区", "医疗事业部"]
        self.department_stats = {}
        self.city_stats = {}
        super().__init__()  # 这行代码调用了AES的构造函数

    def calculate(self):
        for i in range(1, 9):
            self._calculate_department_stats(i)
            self._calculate_city_stats(i)
        return self.department_stats, self.city_stats

    def _calculate_department_stats(self, i):
        for department in self.departments:
            df_department = self._get_df_department(department)
            if department not in self.department_stats:
                self.department_stats[department] = {}
            self.department_stats[department].update(self._get_stats(df_department, i))

    def _get_df_department(self, department):
        if department == "泉州部门":
            quanzhou_departments = ["泉州中小企业增值部", "泉州KOL部门"]
            return self.df[(self.df["部门"].isin(quanzhou_departments)) & (self.df["账户名称"] != "易尔通007")]
        else:
            return self.df[self.df["部门"] == department]

    def _calculate_city_stats(self, i):
        for city in self.df["城市&框架"].unique():
            df_city = self.df[self.df["城市&框架"] == city]
            if city not in self.city_stats:
                self.city_stats[city] = {}
            self.city_stats[city].update(self._get_stats(df_city, i))

    @staticmethod
    def _get_stats(df, i):
        return {
            f"大搜消费第{i}天": df[f"大搜消费第{i}天"].sum(),
            f"信息流消费第{i}天": df[f"信息流消费第{i}天"].sum(),
            f"大搜第{i}天有效消费账数": (df[f"大搜消费第{i}天"] > 0).sum(),
            f"信息流第{i}天有效消费账数": (df[f"信息流消费第{i}天"] > 0).sum(),
            f"易尔通007消费第{i}天": df[df["账户名称"] == "易尔通007"][f"大搜消费第{i}天"].sum(),
            f"易尔通007第{i}天有效消费账数": (df[df["账户名称"] == "易尔通007"][f"大搜消费第{i}天"] > 0).sum(),
            f"feed竞价消费第{i}天": df[f"信息流消费第{i}天"].sum(),
            f"feed竞价账户数第{i}天": (df[f"信息流消费第{i}天"] > 0).sum(),
            f"feed竞价户均第{i}天": df[f"信息流消费第{i}天"].sum() / ((df[f"信息流消费第{i}天"] > 0).sum() or 1),
            f"总计第{i}天": df[f"信息流消费第{i}天"].sum() * 3 / ((df[f"信息流消费第{i}天"] > 0).sum() or 1)
        }

    # 更新缓存
    def update_cache(self, data_1, data_2, data_3):
        dasou = self.calculate_daSou_consumption(data_1, self.load_data_from_cache(self.cache_file_dasou))
        stream = self.calculate_info_stream_consumption(data_2, self.load_data_from_cache(self.cache_file_stream))
        dasou_stream = self.calculate_daSou_info_stream_consumption(data_3,
                                                                    self.load_data_from_cache(
                                                                        self.cache_file_dasou_stream))
        print(dasou)
        print(stream)
        print(dasou_stream)
        return dasou, stream, dasou_stream

    def calculate_daSou_consumption(self, data_dict, cache_data):
        results = {}
        for id, name in data_dict.items():
            customer_data = self.get_customer_data(id)
            if customer_data is not None:
                results[id] = {
                    '客户名称': name,
                    '近七天日均': cache_data.get(id, {}).get('近七天日均', 0),
                    '2023年截止昨日消费': cache_data.get(id, {}).get('2023年截止昨日消费', 0),
                    '前日消费': customer_data['大搜消费第7天'].sum(),
                    '七日均': customer_data['大搜7日均'].sum(),
                    '昨日消费': customer_data['大搜消费第8天'].sum(),
                    '截止消费': cache_data.get(id).get('2023年截止昨日消费') + customer_data[
                        '大搜消费第8天'].sum(),
                    '周一': cache_data.get(id, {}).get('2023年截止昨日消费', 0) + customer_data[
                        '大搜消费第8天'].sum() + customer_data['大搜消费第7天'].sum()
                }

        return results

    def calculate_info_stream_consumption(self, data_dict, cache_data):
        results = {}
        for id, name in data_dict.items():
            customer_data = self.get_customer_data(id)
            if customer_data is not None:
                results[id] = {
                    '客户名称': name,
                    '近七天日均': cache_data.get(id, {}).get('近七天日均', 0),
                    '2023年截止昨日消费': cache_data.get(id, {}).get('2023年截止昨日消费', 0),
                    '前日消费': customer_data['信息流消费第7天'].sum(),
                    '七日均': customer_data['信息流7日均'].sum(),
                    '昨日消费': customer_data['信息流消费第8天'].sum(),
                    '截止消费': cache_data.get(id, {}).get('2023年截止昨日消费', 0) + customer_data[
                        '信息流消费第8天'].sum(),
                    '周一': cache_data.get(id, {}).get('2023年截止昨日消费', 0) + customer_data[
                        '信息流消费第8天'].sum() +
                            customer_data['信息流消费第7天'].sum()
                }
        return results

    def calculate_daSou_info_stream_consumption(self, data_dict, cache_data):
        results = {}
        for id, name in data_dict.items():
            customer_data = self.get_customer_data(id)
            if customer_data is not None:
                results[id] = {
                    '客户名称': name,
                    '大搜日均消费': cache_data.get(id, {}).get('大搜日均消费', 0),
                    '信息流日均消费': cache_data.get(id, {}).get('信息流日均消费', 0),
                    '大搜2023年截止昨日消费': cache_data.get(id, {}).get('大搜2023年截止昨日消费', 0),
                    '信息流2023年截止昨日消费': cache_data.get(id, {}).get('信息流2023年截止昨日消费', 0),
                    '大搜+信息流2023年截止昨日消费': cache_data.get(id, {}).get('大搜+信息流2023年截止昨日消费', 0),
                    '大搜前日消费': customer_data['大搜消费第7天'].sum(),
                    '信息流前日消费': customer_data['信息流消费第7天'].sum(),
                    '大搜+信息流前日消费': customer_data['大搜消费第7天'].sum() + customer_data[
                        '信息流消费第7天'].sum(),
                    '大搜七日均': customer_data['大搜7日均'].sum(),
                    '信息流七日均': customer_data['信息流7日均'].sum(),
                    '大搜昨日消费': customer_data['大搜消费第8天'].sum(),
                    '信息流昨日消费': customer_data['信息流消费第8天'].sum(),
                    '大搜+信息流昨日消费': customer_data['大搜消费第8天'].sum() + customer_data[
                        '信息流消费第8天'].sum(),
                    '大搜截止消费': cache_data.get(id, {}).get('大搜2023年截止昨日消费', 0) + customer_data[
                        '大搜消费第8天'].sum(),
                    '信息流截止消费': cache_data.get(id, {}).get('信息流2023年截止昨日消费', 0) + customer_data[
                        '信息流消费第8天'].sum(),
                    '大搜+信息流截止消费': cache_data.get(id, {}).get('大搜+信息流2023年截止昨日消费', 0) +
                                           customer_data[
                                               '大搜消费第8天'].sum() + customer_data[
                                               '信息流消费第8天'].sum(),
                    '周一大搜截止消费': cache_data.get(id, {}).get('大搜2023年截止昨日消费', 0) + customer_data[
                        '大搜消费第8天'].sum() + customer_data['大搜消费第7天'].sum(),
                    '周一信息流截止消费': cache_data.get(id, {}).get('信息流2023年截止昨日消费', 0) + customer_data[
                        '信息流消费第8天'].sum() + customer_data['信息流消费第7天'].sum(),
                    '周一大搜+信息流截止消费': cache_data.get(id, {}).get('大搜+信息流2023年截止昨日消费', 0) +
                                               customer_data['大搜消费第8天'].sum() + customer_data[
                                                   '大搜消费第7天'].sum() +
                                               customer_data['信息流消费第8天'].sum() + customer_data[
                                                   '信息流消费第7天'].sum(),
                }

        return results

    def get_customer_data(self, id):
        row = self.df[self.df['资质客户ID'] == int(id)]
        if row.empty:
            return None
        else:
            return row


data_1 = {
    "428666192": "厦门快乐番薯股份有限公司",
    "428796862": "厦门运友供应链管理有限公司",
    "429188127": "厦门迪超物流有限公司",
    "429322223": "厦门货运力科技有限公司",
    "429413235": "厦门货小运科技有限公司",
    "429412124": "厦门市湖里区万线帮货运代理服务部",
    "427896986": "厦门雷霆网络科技股份有限公司",
    "56259419": "厦门雷霆互动网络有限公司",
    "428789730": "稿定（厦门）科技有限公司",
    "429542925": "厦门零一世界科技有限公司",
    "428948483": "厦门高定供应链管理有限公司",
    "428344199": "厦门创艺社科技有限公司",
    "428984064": "厦门创艺社管理咨询合伙企业(有限合伙)",
    "428319286": "厦门立马耀网络科技有限公司",
    "429305663": "厦门蝉羽网络科技有限公司",
    "429444813": "厦门蝉客网络科技有限公司",
    "429457902": "厦门康强人才服务有限公司"
}

data_2 = {
    "428666192": "厦门快乐番薯股份有限公司",
    "428987344": "厦门无忧无虑网络科技有限公司",
    "429457902": "厦门康强人才服务有限公司"
}

data_3 = {
    "429334700": "福建朗盛管业科技有限公司",
    "429345878": "福建闽杰管业科技股份有限公司",
    "428007584": "泉州市青果网络科技有限公司",
    "428626133": "厦门房在线科技有限公司",
    "200035639": "厦门快快网络科技有限公司",
    "51611273": "厦门天锐科技股份有限公司",
    "51406109": "厦门市盈拓商务有限公司",
    "428711524": "舒华体育股份有限公司",
    "428953146": "泉州市康掌柜网络科技有限公司",
    "428593822": "福建迈格林医疗科技有限公司",
    "428572382": "福建华雄投资有限公司",
    "55374091": "泉州市丰泽区维美美容美发职业培训学校",
    "429124863": "福建惠兴涂料科技发展有限公司",
    "57373957": "福建惠安县惠兴工贸有限公司",
    "429152891": "福建南方路面机械股份有限公司"
}

stream = {
    "428666192": {
        "客户名称": "厦门快乐番薯股份有限公司",
        "近七天日均": 765,
        "2023年截止昨日消费": 711036,
        "前日消费": 456,
        "七日均": 659,
        "昨日消费": 410,
        "截止消费": 711447,
        "周一": 711902
    },
    "428987344": {
        "客户名称": "厦门无忧无虑网络科技有限公司",
        "近七天日均": 3095,
        "2023年截止昨日消费": 561346,
        "前日消费": 2716,
        "七日均": 2768,
        "昨日消费": 1219,
        "截止消费": 562565,
        "周一": 565281
    },
    "429457902": {
        "客户名称": "厦门康强人才服务有限公司",
        "近七天日均": 4636,
        "2023年截止昨日消费": 958774,
        "前日消费": 4805,
        "七日均": 4589,
        "昨日消费": 4211,
        "截止消费": 962985,
        "周一": 967790
    }
}

dasou_stream = {
    "429334700": {
        "客户名称": "福建朗盛管业科技有限公司",
        "大搜日均消费": 0,
        "信息流日均消费": 0,
        "大搜2023年截止昨日消费": 354602,
        "信息流2023年截止昨日消费": 17000,
        "大搜+信息流2023年截止昨日消费": 371602,
        "大搜前日消费": 0,
        "信息流前日消费": 0,
        "大搜+信息流前日消费": 0,
        "大搜七日均": 0,
        "信息流七日均": 0,
        "大搜昨日消费": 0,
        "信息流昨日消费": 0,
        "大搜+信息流昨日消费": 0,
        "大搜截止消费": 354602,
        "信息流截止消费": 17000,
        "大搜+信息流截止消费": 371602,
        "周一大搜截止消费": 354602,
        "周一信息流截止消费": 17000,
        "周一大搜+信息流截止消费": 371602
    },
    "429345878": {
        "客户名称": "福建闽杰管业科技股份有限公司",
        "大搜日均消费": 1129,
        "信息流日均消费": 0,
        "大搜2023年截止昨日消费": 616174,
        "信息流2023年截止昨日消费": 691,
        "大搜+信息流2023年截止昨日消费": 616866,
        "大搜前日消费": 1173,
        "信息流前日消费": 0,
        "大搜+信息流前日消费": 1173,
        "大搜七日均": 1131,
        "信息流七日均": 0,
        "大搜昨日消费": 1115,
        "信息流昨日消费": 0,
        "大搜+信息流昨日消费": 1115,
        "大搜截止消费": 617289,
        "信息流截止消费": 691,
        "大搜+信息流截止消费": 617981,
        "周一大搜截止消费": 618462,
        "周一信息流截止消费": 691,
        "周一大搜+信息流截止消费": 619153
    },
    "428007584": {
        "客户名称": "泉州市青果网络科技有限公司",
        "大搜日均消费": 2407,
        "信息流日均消费": 0,
        "大搜2023年截止昨日消费": 618922,
        "信息流2023年截止昨日消费": 0,
        "大搜+信息流2023年截止昨日消费": 618922,
        "大搜前日消费": 3371,
        "信息流前日消费": 0,
        "大搜+信息流前日消费": 3371,
        "大搜七日均": 2405,
        "信息流七日均": 0,
        "大搜昨日消费": 3043,
        "信息流昨日消费": 0,
        "大搜+信息流昨日消费": 3043,
        "大搜截止消费": 621965,
        "信息流截止消费": 0,
        "大搜+信息流截止消费": 621965,
        "周一大搜截止消费": 625335,
        "周一信息流截止消费": 0,
        "周一大搜+信息流截止消费": 625335
    }
}

dasou = {
    "428666192": {
        "客户名称": "厦门快乐番薯股份有限公司",
        "近七天日均": 22190,
        "2023年截止昨日消费": 931442,
        "前日消费": 28764,
        "七日均": 22416,
        "昨日消费": 24499,
        "截止消费": 955941,
        "周一": 984705
    },
    "428796862": {
        "客户名称": "厦门运友供应链管理有限公司",
        "近七天日均": 2729,
        "2023年截止昨日消费": 459616,
        "前日消费": 2452,
        "七日均": 2624,
        "昨日消费": 2282,
        "截止消费": 461899,
        "周一": 464351
    },
    "429188127": {
        "客户名称": "厦门迪超物流有限公司",
        "近七天日均": 765,
        "2023年截止昨日消费": 188543,
        "前日消费": 644,
        "七日均": 731,
        "昨日消费": 622,
        "截止消费": 189165,
        "周一": 189809
    },
    "429322223": {
        "客户名称": "厦门货运力科技有限公司",
        "近七天日均": 580,
        "2023年截止昨日消费": 158509,
        "前日消费": 465,
        "七日均": 559,
        "昨日消费": 423,
        "截止消费": 158932,
        "周一": 159397
    },
    "429413235": {
        "客户名称": "厦门货小运科技有限公司",
        "近七天日均": 499,
        "2023年截止昨日消费": 61338,
        "前日消费": 474,
        "七日均": 481,
        "昨日消费": 440,
        "截止消费": 61778,
        "周一": 62252
    },
    "429412124": {
        "客户名称": "厦门市湖里区万线帮货运代理服务部",
        "近七天日均": 0,
        "2023年截止昨日消费": 0,
        "前日消费": 0,
        "七日均": 0,
        "昨日消费": 0,
        "截止消费": 0,
        "周一": 0
    },
    "427896986": {
        "客户名称": "厦门雷霆网络科技股份有限公司",
        "近七天日均": 16669,
        "2023年截止昨日消费": 3308133,
        "前日消费": 18244,
        "七日均": 17318,
        "昨日消费": 22108,
        "截止消费": 3330241,
        "周一": 3348485
    },
    "56259419": {
        "客户名称": "厦门雷霆互动网络有限公司",
        "近七天日均": 0,
        "2023年截止昨日消费": 0,
        "前日消费": 0,
        "七日均": 0,
        "昨日消费": 0,
        "截止消费": 0,
        "周一": 0
    },
    "428789730": {
        "客户名称": "稿定（厦门）科技有限公司",
        "近七天日均": 19706,
        "2023年截止昨日消费": 7513661,
        "前日消费": 25197,
        "七日均": 19953,
        "昨日消费": 23942,
        "截止消费": 7537604,
        "周一": 7562800
    },
    "429542925": {
        "客户名称": "厦门零一世界科技有限公司",
        "近七天日均": 1485,
        "2023年截止昨日消费": 396013,
        "前日消费": 1486,
        "七日均": 1488,
        "昨日消费": 1277,
        "截止消费": 397290,
        "周一": 398776
    },
    "428948483": {
        "客户名称": "厦门高定供应链管理有限公司",
        "近七天日均": 727,
        "2023年截止昨日消费": 119674,
        "前日消费": 943,
        "七日均": 776,
        "昨日消费": 1064,
        "截止消费": 120738,
        "周一": 121681
    },
    "428344199": {
        "客户名称": "厦门创艺社科技有限公司",
        "近七天日均": 2847,
        "2023年截止昨日消费": 470057,
        "前日消费": 2533,
        "七日均": 2792,
        "昨日消费": 3275,
        "截止消费": 473332,
        "周一": 475865
    },
    "428984064": {
        "客户名称": "厦门创艺社管理咨询合伙企业（有限合伙）",
        "近七天日均": 303,
        "2023年截止昨日消费": 58149,
        "前日消费": 347,
        "七日均": 303,
        "昨日消费": 366,
        "截止消费": 58514,
        "周一": 58861
    },
    "428319286": {
        "客户名称": "厦门立马耀网络科技有限公司",
        "近七天日均": 0,
        "2023年截止昨日消费": 0,
        "前日消费": 0,
        "七日均": 0,
        "昨日消费": 0,
        "截止消费": 0,
        "周一": 0
    },
    "429305663": {
        "客户名称": "厦门蝉羽网络科技有限公司",
        "近七天日均": 7519,
        "2023年截止昨日消费": 1021467,
        "前日消费": 8303,
        "七日均": 7582,
        "昨日消费": 8331,
        "截止消费": 1029798,
        "周一": 1038101
    },
    "429444813": {
        "客户名称": "厦门蝉客网络科技有限公司",
        "近七天日均": 0,
        "2023年截止昨日消费": 0,
        "前日消费": 0,
        "七日均": 0,
        "昨日消费": 0,
        "截止消费": 0,
        "周一": 0
    },
    "429457902": {
        "客户名称": "厦门康强人才服务有限公司",
        "近七天日均": 7590,
        "2023年截止昨日消费": 1444136,
        "前日消费": 7785,
        "七日均": 7668,
        "昨日消费": 7927,
        "截止消费": 1452063,
        "周一": 1459848
    }
}
