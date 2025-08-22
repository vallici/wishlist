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
        <title>Clara's Wishlist</title>
        <style>
            body {{
            
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 20px;
				font-weight: bold
                
            }}
            
            h1 {{
            
                text-align: center;
                
            }}
            
            .wishlist-container {{
                
				display: table;
				margin: 0 auto;
                max-width: 70em;
                width: 100%;
                height: auto;
                position: relative;
                overflow: hidden;
                border-spacing: 0 1em;

                
            }}
            
            .wishlist-item {{
            
				display: table-row;
                background-color: #f9f9f9;
                
            }}
			

            
            .wishlist-item:nth-child(even) {{
            
                background-color: #e9e9e9;
                
            }}
            
			.wishlist-image {{
            
                display:table-cell;
				width: 15em;
                height: 10em;
				min-witdh: 15em;
				vertical-align: middle;
				position: relative;
                align-items: center;
                text-align: center;
                
			}}
			
            .wishlist-image img {{
            
				display:block;
				max-width: 15em;
                max-height: 10em;
                margin-right: auto;
				margin-left: auto;
                width: auto;
                height: auto;
				vertical-align:middle;
                color: #F527A3;

            }}
			
           .wishlist-name {{
           
				display:table-cell;
				text-decoration: none;
                color: #000;
				min-width: 10em;
                max-width: 40em;
				align-items:center;
				margin-right: 2em;
				vertical-align:middle;
				margin-left: 2em;
				padding-left: 2em;

            }}
            .wishlist-item a {{
            
                text-decoration: none;
                color: #000;
                
             }}
            
            .wishlist-price  {{
            
				display:table-cell;
				vertical-align:middle;
				align-items:center;
				text-align: center-right;
				padding-left: 1em;
				margin-left: 1em;
                width: auto;

            }}
            
        </style>
    </head>
    <body>
        <h1>{sheet_name}</h1>
        
        <div class="wishlist-container">
        """.format(sheet_name=sheet_name)

    for index, item in enumerate(items):
        if item['reserved'] == 'Y':
            continue  # Skip reserved items
        
        html += """
            <div class="wishlist-item">
                <div class="wishlist-image">"""
        if item['image']:
            html += """
                <img src="{item_image}" alt="missing image">""".format(
                    item_image=item['image']
                    )
        else:
            html += """
                n/a"""
        html += """
            </div>
            <div class="wishlist-name">"""
        if item['link'] and item['name']:
            html += """
                <a href="{item_link}" target="_blank">
                    {item_name}
                </a>""".format(
                    item_link = item['link'],
                    item_name = item['name']
                    )
        else:
            if item['link']:
                html += """
                    <a href="{item_link}" target="_blank" >
                        {item_link}
                    </a>""".format(
                    item_link = item['link'],
                    )
            else:
                html += """
                    {item_name} 
                    <span style="color: #F527A3">
                    (no link)
                        </span>""".format(
                    item_name = item['name'],
                    )
        html += """
            </div>
            <div class="wishlist-price">"""
        if item['price']:
            html += """
                {item_price}""".format(
                    item_price = item['price']
                    )
        else:
            html += """
                n/a"""
        html += """
            </div></div>"""
        

    html += """
        </div>
    </body>
    </html>"""
    
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
