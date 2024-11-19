import openpyxl
from openpyxl.comments import Comment
from html import escape

def excel_to_html_with_comments(excel_file, html_file):
    # Load the workbook
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active

    # Start the HTML content
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Excel Sheet with Comments</title>
        <style>
            table { border-collapse: collapse; }
            th, td { border: 1px solid black; padding: 8px; position: relative; }
            .comment { display: none; position: absolute; background: #ffffe0; border: 1px solid #000; padding: 5px; z-index: 100; }
            td:hover .comment { display: block; }
        </style>
    </head>
    <body>
        <table>
    """

    # Iterate through rows and columns
    for row in sheet.iter_rows():
        html_content += "<tr>"
        for cell in row:
            cell_value = escape(str(cell.value)) if cell.value is not None else ""
            if cell.comment:
                comment_text = escape(cell.comment.text)
                html_content += f'<td>{cell_value}<div class="comment">{comment_text}</div></td>'
            else:
                html_content += f'<td>{cell_value}</td>'
        html_content += "</tr>"

    # Close the HTML content
    html_content += """
        </table>
    </body>
    </html>
    """

    # Write the HTML content to a file
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(html_content)

# Usage
excel_to_html_with_comments('feature-comparison.xlsx', 'table.html')