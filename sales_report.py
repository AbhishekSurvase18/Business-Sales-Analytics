import pandas as pd

df = pd.read_csv("sales.csv")

df.columns = df.columns.str.strip()

df["Total"] = df["Price"] * df["Quantity"]

total_sales = df["Total"].sum()
average_sale = df["Total"].mean()
highest_sale = df.loc[df["Total"].idxmax()]

category_sales = (
    df.groupby("Category")["Total"]
    .sum()
    .reset_index()
)

with pd.ExcelWriter(
    "Business_Sales_Report.xlsx",
    engine="xlsxwriter"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Sales Data",
        index=False
    )

    workbook = writer.book

    header_format = workbook.add_format({
        "bold": True,
        "border": 1,
        "align": "center"
    })

    currency_format = workbook.add_format({
        "num_format": "₹#,##0.00"
    })

    sales_sheet = writer.sheets["Sales Data"]

    for col_num, value in enumerate(df.columns):
        sales_sheet.write(
            0, col_num, value, header_format
        )

    sales_sheet.set_column("A:A", 18)
    sales_sheet.set_column("B:B", 15, currency_format)
    sales_sheet.set_column("C:C", 12)
    sales_sheet.set_column("D:D", 18)
    sales_sheet.set_column("E:E", 15, currency_format)

    summary = pd.DataFrame({
        "Metric": [
            "Total Sales",
            "Average Sale",
            "Highest Sale Product",
            "Highest Sale Amount"
        ],
        "Value": [
            total_sales,
            average_sale,
            highest_sale["Product"],
            highest_sale["Total"]
        ]
    })

    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    summary_sheet = writer.sheets["Summary"]

    summary_sheet.set_column("A:A", 25)
    summary_sheet.set_column("B:B", 25)

    category_sales.to_excel(
        writer,
        sheet_name="Category Sales",
        index=False
    )

    category_sheet = writer.sheets["Category Sales"]

    category_sheet.set_column("A:A", 20)
    category_sheet.set_column(
        "B:B", 20, currency_format
    )

    chart = workbook.add_chart({
        "type": "column"
    })

    chart.add_series({
        "name": "Category Sales",
        "categories": "='Category Sales'!$A$2:$A$10",
        "values": "='Category Sales'!$B$2:$B$10"
    })

    chart.set_title({
        "name": "Category-wise Sales"
    })

    chart.set_x_axis({
        "name": "Category"
    })

    chart.set_y_axis({
        "name": "Sales Amount"
    })

    chart.set_legend({
        "none": True
    })

    category_sheet.insert_chart(
        "D2",
        chart,
        {"x_scale": 1.5, "y_scale": 1.5}
    )

print("================================")
print(" Professional Sales Report")
print("================================")
print(f"Total Sales: ₹{total_sales:,.2f}")
print(
    f"Highest Sale: {highest_sale['Product']} "
    f"- ₹{highest_sale['Total']:,.2f}"
)
print(f"Average Sale: ₹{average_sale:,.2f}")
print("Excel report created successfully!")
print("File: Business_Sales_Report.xlsx")