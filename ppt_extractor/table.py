import pandas as pd

def extract_table(shape):
    rows_data = []
    for row in shape.table.rows:
        row_data = []
        for cell in row.cells:
            row_data.append(cell.text)
        rows_data.append(row_data)
    df = pd.DataFrame(rows_data)

    # Code for having first row as a column
    # df.columns = df.iloc[0]
    # df = df[1:]

    table_data = "The following is a table data with default columns : " + "\n"
    for index, row in df.iterrows():
        for index, item in row.items():
            table_data += f"{index} : {item}"
            table_data += ", "
        table_data += "\n"
    
    data = {
        "data" : table_data,
        "shape_type" : shape.shape_type
    }
    return data