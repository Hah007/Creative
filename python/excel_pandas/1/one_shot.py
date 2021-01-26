
#%%
import xlwings as xw
import pandas as pd
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False
wb=app.books.open('成绩单.xlsx')
wb=xw.books['成绩单.xlsx']
wrk=wb.sheets['Sheet1']

df=wrk.range('a1').current_region.options(pd.DataFrame).value
# 透视
print(df)
print("***************")
pv_df=pd.pivot_table(df,
        index='班级',
        margins=True,
        margins_name='科目总平均')
print(pv_df)
# 调整结果
cols='语文,数学,英语,物理,历史,地理,政治,生物,总分'.split(',')
pv_df=pv_df[cols]
pv_df.reset_index(inplace=True)
print("&"*20)
print(pv_df)
# 输出
wrk.range('N1').value='苦短'
wrk.range('O11').value=pv_df.columns.tolist()
wrk.range('O12').value=pv_df.values
print(df)
wb.save()
wb.close()
app.quit()
#%%
df

#%%
