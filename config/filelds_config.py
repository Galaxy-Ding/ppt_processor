# 定义每个需求要提取那些页面/字段
# src/office_ops/ppt_processor/config/fields_config.py
FIELDS_CONFIG = {
    "发包规范V1": {
        "master":{
            "iou":{
                "ProjectCode": (22.17, 1.33, 3.54, 2.11),
            }
        },
        "page":{
            "1":{
                "iou": {
                    "DFM_Info": {
                        "box": (14.44, 15.58, 4.15, 3.23), # 制作人信息，制作时间，制作单位，联系方式 eg: '方案中心\n李爱民\n568+42+123456\n2024-01-01'
                        "need_split": "\n",  # 采用的分割的字符
                        "storage_var": ["", "engineer", "", "changing_date"], # 分割的字符，存储在那个字段，例如，list ： 0-为空，则不进行存储，1不为空，则对应的也是分割后的索引进行存储。
                    },
                }
            }

        },
        "title":[
            {
            "first": "設備規格及參數", # 1 级
            "second": "", # 2 级
            "re": {
                # re 匹配文字字段 并且提取对应的数值
                # 每个的key 是存储的变量名称
                "dev_max_size": {
                    "match_key_string": "場地佔用", # 匹配的字符串
                    "re_rule": r"[場场]地[佔占]用.*[:：].*mm",
                    "match_rule": -1,  # 匹配规则: -1 ： 表示需要做对应的中间字符作提取 0：表示当前的文本框；1：找寻最近图片类 或者 custom shape 方向默认 下方 ，2 则是左侧，3则是上方，4则是右边
                },
                "dev_failure_rate": {
                    "match_key_string": "故障率", # 匹配的字符串
                    "re_rule": r"故障率\s*[:：]\s*(.*?)(?=\n\d+\.|\n\s*$|\n\s*\d+[\.。]|\Z)",
                    "match_rule": -1,  # 匹配规则: -1 ： 表示需要做对应的中间字符作提取0：表示当前的文本框；1：找寻最近图片类 或者 custom shape 方向默认 下方 ，2 则是左侧，3则是上方，4则是右边
                },
                "ComprehensiveCT": {
                    "match_key_string": "綜合CT", # 匹配的字符串
                    "re_rule": r"[綜综]合[cC][Tt].*[:：]\s*([\d\.]+[Ss]/[Pp][Cc][Ss].*?)\s",
                    "match_rule": -1,  # 匹配规则: -1 ： 表示需要做对应的中间字符作提取0：表示当前的文本框；1：找寻最近图片类 或者 custom shape 方向默认 下方 ，2 则是左侧，3则是上方，4则是右边
                },
                "ComprehensiveUPH": {
                    "match_key_string": "UPH", # 匹配的字符串
                    "re_rule": r"[Uu][Pp][Hh].*[:：].*([Pp][Cc][Ss]/[Hh])",
                    "match_rule": -1,  # 匹配规则: -1 ： 表示需要做对应的中间字符作提取0：表示当前的文本框；1：找寻最近图片类 或者 custom shape 方向默认 下方 ，2 则是左侧，3则是上方，4则是右边
                },
                "dev_overkill_rate": {
                    "match_key_string": "過殺率", 
                    "re_rule": r"[過过][殺杀]率.*[(（][檢检][測测][設设][備备][)）][:：]\s*([/\d\.%]+)",
                    "match_rule": -1,  # 匹配规则: -1 ： 表示需要做对应的中间字符作提取0：表示当前的文本框；1：找寻最近图片类 或者 custom shape 方向默认 下方 ，2 则是左侧，3则是上方，4则是右边
                },
                "dev_miss_rate": {
                    "match_key_string": "漏檢率", 
                    "re_rule": r"漏[檢检]率.*[(（][檢检][測测][設设][備备][)）][:：]\s*([/\d\.%]+)",
                    "match_rule": -1,  # 匹配规则: -1 ： 表示需要做对应的中间字符作提取0：表示当前的文本框；1：找寻最近图片类 或者 custom shape 方向默认 下方 ，2 则是左侧，3则是上方，4则是右边
                },
                "dev_operation_manpower": {
                    "match_key_string": "機台操作人力", 
                    "re_rule": r"[机機]台操作人力\s*[:：].*人/[机機]",
                    "match_rule": -1,  # 匹配规则: -1 ： 表示需要做对应的中间字符作提取0：表示当前的文本框；1：找寻最近图片类 或者 custom shape 方向默认 下方 ，2 则是左侧，3则是上方，4则是右边
                }
            }
        },
        {
            "first": "二.改造方案介紹及模組說明", # 1 级
            "second": "1.方案整體概況", # 2 级
            "re": {
                # re 匹配文字字段：
                "img_dev_frontlooking": { # 图片
                    "match_key_string": "4.長寬高尺寸,佔地面積",
                    "re_rule": r"(?:\d+[\.。]\s*)?[长長][宽寬]高尺寸\s*[，,]\s*[佔占]地面[積积][:：]",
                    "match_rule": 1,  # 匹配规则: 0：表示当前的文本框；1：找寻最近图片类 或者 custom shape 方向默认 下方 ，2 则是左侧，3则是上方，4则是右边
                },
                "img_dev_occupancy": { # 文本框本身
                    "match_key_string": "4.長寬高尺寸,佔地面積",
                    "re_rule": r"(?:\d+[\.。]\s*)?[长長][宽寬]高尺寸\s*[，,]\s*[佔占]地面[積积][:：]",
                    "match_rule": 5,  # 匹配规则: 0：表示当前的文本框；1：找寻最近图片类 或者 custom shape 方向默认 下方 ，2 则是左侧，3则是上方，4则是右边，5 则是本身，需要被当作图被返回
                },
                "img_dev_overlooking": { # 图片
                    "match_key_string": "3.俯視佈局圖",
                    "re_rule": r"(?:\d+[\.。]\s*)?俯[视視][布佈]局[图圖][:：]",
                    "match_rule": 1,  # 匹配规则: 0：表示当前的文本框；1：找寻最近图片类 或者 custom shape 方向默认 下方 ，2 则是左侧，3则是上方，4则是右边
                }
            }
        },
        {
            "first": "二.改造方案介紹及模組說明", # 1 级
            "second": "2.工藝流程", # 2 级
            "re": {
                # re 匹配文字字段：
                "img_dev_craftsmanship": { # 图片工艺流程的图片
                    "match_key_string": "工藝流程介紹",
                    "re_rule": r"^\s*(\d+\s*\.?\s*)?\s*工\s*[艺藝]\s*流\s*程\s*介\s*[绍紹]",
                    "match_rule": 1,  # 匹配规则: 0：表示当前的文本框；1：找寻最近图片类 或者 custom shape 方向默认 下方 ，2 则是左侧，3则是上方，4则是右边
                }
            }
        },
        {
            "first": "方案版本變更記錄", # 1 级
            "second": "", # 2 级
            "table": {
                # 默认取其最后一行
                "version": {
                    "match_key_string": "報告版本\n（版本號+報告日期）",
                },
                "changing_description": {
                    "match_key_string": "變更內容\n（需說明變更前和變更后內容對比）",
                },
                "release_date": {
                    "match_key_string": "變更\n日期",
                }
            }
        }

        ]
    }
}