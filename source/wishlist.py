import openpyxl
import requests
import sys
from bs4 import BeautifulSoup




def generate_wishlist_html_from_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    active_sheet = workbook.active
    sheet_name = active_sheet.title

    items = []
    
    for row in active_sheet.iter_rows(min_row=2, values_only=True):
        item = {
            'name': row[0],
            'image': row[1],
            'link': row[2],
            'price': row[3],
            'reserved': row[4]
        }
        items.append(item)

    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>{sheet_name}'s Wishlist</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 20px;
            }}
            
            h1 {{
                text-align: center;
            }}
            
            .wishlist-container {{
                margin: 0 auto;
                max-width: 70em;
                border: 1px solid #ccc;
                padding: 1em;
            }}
            
            .wishlist-item {{
                margin-bottom: 1em;
                display: flex;
                align-items: center;
                background-color: #f9f9f9;
            }}
            
            .wishlist-item:nth-child(even) {{
                background-color: #e9e9e9;
            }}
            
            .wishlist-item img {{
                display: block;
                max-width: 15em;
                margin-right: 1em;
            }}
            
            .wishlist-item a {{
                text-decoration: none;
                color: #000;
                font-weight: bold;
                width: 40em;
            }}
            
            .wishlist-item .price {{
                font-weight: bold;
                text-align: right;
                padding-left: 1em;
                padding-right: 1em;
                width: 8em;
            }}
        </style>
    </head>
    <body>
        <h1>{sheet_name}'s Wishlist</h1>
        
        <div class="wishlist-container">
    """.format(sheet_name=sheet_name)

    for index, item in enumerate(items):
        if item['reserved'] == 'Y':
            continue  # Skip reserved items
        
        if item['name']:
            item_name = item['name']
        else:
#           item_name = get_page_title(item['link'])
            item_name = item['link']
        
        html += """
            <div class="wishlist-item">
                <img src="{item_image}" alt="wishlist-item">
                <a href="{item_link}" target="_blank">{item_name}</a>
                <div class="price">{item_price}</div>
            </div>
        """.format(
            item_image=item['image'],
            item_link=item['link'],
            item_name=item_name,
            item_price=item['price']
        )
        
    html += """
        </div>
    </body>
    </html>
    """

    return html


#def get_page_title(url):
#    response = requests.get(url)
#    soup = BeautifulSoup(response.text, 'html.parser')
#    page_title = soup.title.string.strip()
#    return page_title


# Example usage
file_path = sys.argv[1]
wishlist_html = generate_wishlist_html_from_excel(file_path)
print(wishlist_html)
