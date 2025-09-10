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

    html = """<!DOCTYPE html>
    <html>
    <head>
        <title>{sheet_name}</title>
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

    ### Set an error color
    
    error_color = "#F527A3"
    normal_color = "#000000"
    
    ### Set opacity levels
    reserved_opacity = "0.3"
    normal_opacity = "1"
    
    ### Set height levels
    gap_height = "4em"
    normal_height = "10em"
    
    ### Set an error text
    na_text = "n/a"
    na_formatted_text = "<span style = \"color: " + error_color + "\">" + na_text + "</span>"
    
    ### Set the extra backgroud color modifier
    
    bkg_color = " background-color: #ccbddb;"
      
    for index, item in enumerate(items):
        
        ### Check if it needs to skip the item
        if item['reserved'] == "SKIP":
            continue
        
        if not item['name'] and  not item['link'] and  not item['image'] and not item['reserved']:
            continue
        
        ### Set the background color to default before evaluating if this is a gap
        html_bkg = ""
        
        ### Check items in the list
        
        ### Check if the image exists
        if item['image']:
            html_image = """<img src=""" + item['image'] + """ alt="missing image">"""
        else:
            html_image = na_formatted_text
        
        ### Check if the price exists
        if item['price']:
            html_price = item['price']
        else:
            html_price = na_formatted_text
        
        ### Check if the link and the name both exist
        if item['link'] and item['name']:
            html_name = """<a href=""" + item['link'] + """ target="_blank"> """ + item['name'] + "</a>"
        else:
            ### If just the link is there, but no item name
            if item['link']:
                html_name = """<a href=""" + item['link'] + """ target="_blank"></a>"""
            else:
                if item['name']:
                    ### If just the item name is there, but no item link
                    html_name = item['name'] + """<span style="color: {error_color}"> (no link) </span>""".format(error_color = error_color)
                else:
                    html_name = na_formatted_text
                    
        ### Check if the item is reserved or marked as a gap
        if item['reserved'] == "Y":
            html_opacity = reserved_opacity
            html_height = normal_height

        else:
            html_opacity = normal_opacity
            html_height = normal_height
        
        if item['reserved'] == "GAP":
            html_bkg = bkg_color
            html_height = gap_height
            html_price = ""
            html_image = ""
            html_name = ""
                 
        html += """
            <div class="wishlist-item" style="opacity: {html_opacity}; height: {html_height};{html_bkg}">
                <div class="wishlist-image">
                    
                    {html_image}

                </div>
                <div class="wishlist-name">
                
                    {html_name}
                
                </div>
                <div class="wishlist-price">
                
                    {html_price}
                
                </div>
            </div>""".format(
                        html_bkg = html_bkg,
                        html_opacity = html_opacity,
                        html_height = html_height,
                        html_image = html_image,
                        html_name = html_name,
                        html_price = html_price
                        )
   

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
