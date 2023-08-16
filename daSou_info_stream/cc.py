# Step 1: Read the file and preprocess the data

file_path = '/mnt/data/新建 文本文档 (2).txt'

with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

cleaned_data = []
for line in lines:
    # remove the newline character and comma
    line = line.replace('\n', '').replace(',', '')
    # split the line into list of strings
    cleaned_data.append(line.split())

# Step 2: Convert the preprocessed data into dictionary

dasou_stream = {}
for line in cleaned_data:
    key = line[0]
    values = line[1:]
    dasou_stream[key] = {
        "客户名称": values[0],
        "大搜日均消费": int(values[1]),
        "信息流日均消费": int(values[2]),
        "大搜2023年截止昨日消费": int(values[3]),
        "信息流2023年截止昨日消费": int(values[4]),
        "大搜+信息流2023年截止昨日消费": int(values[5]),
        "大搜前日消费": int(values[6]),
        "信息流前日消费": int(values[7]),
        "大搜+信息流前日消费": int(values[8]),
        "大搜七日均": int(values[9]),
        "信息流七日均": int(values[10]),
        "大搜昨日消费": int(values[11]),
        "信息流昨日消费": int(values[12]),
        "大搜+信息流昨日消费": int(values[13]),
        "大搜截止消费": int(values[14]),
        "信息流截止消费": int(values[15]),
        "大搜+信息流截止消费": int(values[16]),
        "周一大搜截止消费": int(values[17]),
        "周一信息流截止消费": int(values[18]),
        "周一大搜+信息流截止消费": int(values[19])
    }

# Step 3: Write the dictionary into a Python file

python_file_path = '/mnt/data/base.py'
with open(python_file_path, 'w', encoding='utf-8') as f:
    f.write('data_dict = ' + repr(dasou_stream))

