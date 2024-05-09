import pandas as pd

# 读取Excel文件
def read_excel(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
    except ValueError:
        print(f"未找到指定的Sheet页:'{sheet_name}',尝试读取第一个Sheet页。")
        df = pd.read_excel(file_path, sheet_name=0, header=1)
        expected_headers = ['柜员号', '统一认证号', '员工姓名', '柜员状态', '公司', '小组', '入职日期', '是否培训期', '转正日期', '任务类型', '查询日期']
        if not set(expected_headers).issubset(df.columns):
            raise ValueError("第一个Sheet页的内容没有正确的表头。")
    return df

# Task1
def task1(df):
    # 去重并计算上线人数。这里使用drop_duplicates函数去除DataFrame中重复的行，
    # subset参数指定了去重时考虑的列（即根据'统一认证号'和'查询日期'两列去重）。
    # shape[0]用于获取去重后DataFrame的行数，即上线人数。
    online_num = df.drop_duplicates(subset=['统一认证号', '查询日期']).shape[0]
    # 计算有效通话量
    valid_call_volume = df['统一认证号'].count()
    # 输出结果
    result = pd.DataFrame({
        '公司': [df['公司'].iloc[0]],
        '查询日期': [df['查询日期'].iloc[0]],
        '上线人数': [online_num],
        '有效通话量合计': [valid_call_volume]
    })
    return result

# Task2
def task2(df):
    # 计算工作天数。这里使用相同的方法计算'统一认证号'和'查询日期'去重后的记录数，
    # 代表了每个员工的工作天数。
    work_days = df.drop_duplicates(subset=['统一认证号', '查询日期']).shape[0]
    # 计算有效通话量
    valid_call_volume = df['统一认证号'].count()
    # 输出结果
    result = pd.DataFrame({
        '公司': [df['公司'].iloc[0]],
        '统一认证号': [df['统一认证号'].iloc[0]],
        '工作天数': [work_days],
        '有效通话量': [valid_call_volume],
        '月份': [df['查询日期'].iloc[0].month]
    })
    return result

# 主函数
def main():
    try:
        df = read_excel('驻场催收人员明细报表.xlsx', '驻场催收人员明细报表')
        result1 = task1(df)
        result2 = task2(df)
        
        with pd.ExcelWriter('N1结算表.xlsx', engine='openpyxl') as writer:
            result1.to_excel(writer, sheet_name='公司维度', index=False)
            result2.to_excel(writer, sheet_name='坐席维度', index=False)
    except Exception as e:
        print(f"程序出错：{e}")

if __name__ == "__main__":
    main()