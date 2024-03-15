import pandas as pd
from flask import Flask, render_template, request, send_file
import re

app = Flask(__name__)

@app.route('/')
def eazereport():
    return render_template('eazereport.html')


def categorize_Segment(Segment):
    if re.search(r'Added|added', Segment):
        return 'Added'
    elif re.search(r'Active|active', Segment):
        return 'Active'
    elif re.search(r'Clicker|clicker', Segment):
        return 'Clicker'
    else:
        return Segment


@app.route('/process', methods=['POST'])
def process():
    file = request.files['file']
    df = pd.read_excel(file)
    pivot_tables = []
    definetype = int(request.form.get('report_type'))
    if definetype == 1:
        definetype = 'ESP'
        index_type = [definetype,'Segment Category']
    elif definetype == 2:
        definetype = 'List'
        index_type = [definetype,'Segment Category']
    elif definetype == 3:
        definetype = 'Segment'
        index_type = ['Segment Category','List']
    elif definetype == 4:
        definetype = 'ESP'
        index_type = [definetype]
    elif definetype == 5:
        definetype = 'List'
        index_type = [definetype]
    elif definetype == 6:
        definetype = 'onlysegment'
        index_type = ['Segment Category']
    else: 
        return "Kindly select proper option from the list"
        
    df['Segment Category'] = df['Segment'].apply(categorize_Segment)
    pivot_table = pd.pivot_table(
            df,
            index= index_type,
            values=['Count', 'G_OPEN','Open', 'G_CLICK','Clicks','Unsub','Con','Nw Click','Rev','Complaints','Failed'],
            aggfunc='sum')
    pivot_tables.append(pivot_table)
    result_pivot_table = pd.concat(pivot_tables)

    # Reordering columns
    desired_order = ['Count', 'G_OPEN', 'Open', 'G_CLICK', 'Clicks', 'Unsub', 'Con','Nw Click','Rev','Complaints']
    result_pivot_table = result_pivot_table[desired_order]

    # Saving to Excel
    import uuid
    output_file_name = definetype + "_report_"+ str(uuid.uuid4())[:8]  + ".xlsx"
    result_pivot_table.to_excel(output_file_name)
    return send_file(output_file_name, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
