import pandas as pd
#อ่านข้อมูลจากไฟล์ Excel
df = pd.read_excel('shop_orders.xlsx')
#คำนวณยอดขายโดยการคูณราคาและจำนวน
df["ยอดขาย"] = df["ราคา"] * df["จำนวน"]
#คำนวนยอดขายรวม
total_sales = df.groupby("สินค้า")["ยอดขาย"].sum()
#หาสินค้าที่ขายดีที่สุด
max_sales = df.groupby("สินค้า")["ยอดขาย"].sum().idxmax()
#หาวันที่ขายดีที่สุด
best_selling_product = df.groupby("วันที่")["ยอดขาย"].sum()
#หายอดขายที่สูงที่สุด
top = best_selling_product.idxmax()


#สรุปลงในไฟล์ Excel ใหม่
with pd.ExcelWriter('shop_report.xlsx') as writer:
    df.to_excel(writer, sheet_name='Orders', index=False)
    total_sales.to_excel(writer, sheet_name='Monthly Report')
    best_selling_product.to_excel(writer, sheet_name='ยอดขายในแต่ละวัน')
    best_selling_product[[top]].to_excel(writer, sheet_name='วันที่ขายดีที่สุด')
print("รายงานยอดขายถูกบันทึกในไฟล์ shop_report.xlsx")