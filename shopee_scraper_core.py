import re
import requests
import os
import time
import pandas as pd
from typing import Optional, Dict, Any, List, Set

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError

# Google Sheets API Scopes
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def parse_proxy(proxy_str: str) -> Optional[Dict[str, str]]:
    """
    Parse proxy string to requests format.
    Format input: ip:port:user:pass or ip:port
    Format output: {'http': '...', 'https': '...'}
    """
    if not proxy_str or not proxy_str.strip():
        return None
    
    parts = proxy_str.strip().split(':')
    if len(parts) == 4:
        ip, port, user, password = parts
        proxy_url = f"http://{user}:{password}@{ip}:{port}"
        return {"http": proxy_url, "https": proxy_url}
    elif len(parts) == 2:
        ip, port = parts
        proxy_url = f"http://{ip}:{port}"
        return {"http": proxy_url, "https": proxy_url}
    
    return None

def extract_item_id(input_str: str) -> Optional[str]:
    """
    Extract item_id from Shopee product link or ID.
    Supports formats:
    - 54005267037 (pure digits)
    - https://shopee.vn/product/302865535/54005267037
    - https://shopee.vn/Ao-i.302865535.54005267037
    """
    input_str = input_str.strip()
    if not input_str:
        return None
    
    # Check if pure digits
    if input_str.isdigit():
        return input_str
    
    # Try pattern: i.shopid.itemid
    match_i = re.search(r'i\.(\d+)\.(\d+)', input_str)
    if match_i:
        return match_i.group(2)
        
    # Try pattern: product/shopid/itemid
    match_prod = re.search(r'product/(\d+)/(\d+)', input_str)
    if match_prod:
        return match_prod.group(2)
        
    return None

def extract_shop_id_from_link(input_str: str) -> Optional[str]:
    """
    Extract shop_id from Shopee product link if present.
    """
    input_str = input_str.strip()
    if not input_str:
        return None
        
    # Try pattern: i.shopid.itemid
    match_i = re.search(r'i\.(\d+)\.(\d+)', input_str)
    if match_i:
        return match_i.group(1)
        
    # Try pattern: product/shopid/itemid
    match_prod = re.search(r'product/(\d+)/(\d+)', input_str)
    if match_prod:
        return match_prod.group(1)
        
    return None

def fetch_product_data(item_id: str, proxy_str: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """
    Fetch product data from Addlivetag API.
    Retries up to 3 times on network failure.
    """
    url = f"https://data.addlivetag.com/product-data/product-data.php?item_id={item_id}"
    proxies = parse_proxy(proxy_str) if proxy_str else None
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json"
    }
    
    retries = 3
    for attempt in range(1, retries + 1):
        try:
            response = requests.get(url, headers=headers, proxies=proxies, timeout=15)
            if response.status_code == 200:
                data = response.json()
                if data.get("status") == "success" and "productInfo" in data:
                    info = data["productInfo"]
                    
                    # Extract shopId from productLink if it exists
                    prod_link = info.get("productLink", "")
                    shop_id = extract_shop_id_from_link(prod_link) or ""
                    
                    return {
                        "Mã SP (Item ID)": str(info.get("itemId", item_id)),
                        "Mã Shop (Shop ID)": str(shop_id),
                        "Tên sản phẩm": info.get("productName", "N/A"),
                        "Giá (đ)": int(info.get("price", 0)),
                        "Lượt bán": int(info.get("sales", 0)),
                        "Đường dẫn sản phẩm": prod_link,
                        "Hoa hồng (đ)": int(info.get("commission", 0)),
                        "Hoa hồng người bán (đ)": int(info.get("sellerComFinal", 0)),
                        "Hoa hồng Shopee (đ)": int(info.get("shopeeComFinal", 0))
                    }
                else:
                    raise ValueError("API status is not success or productInfo missing.")
            else:
                raise requests.HTTPError(f"HTTP Status {response.status_code}")
                
        except Exception as e:
            if attempt < retries:
                time.sleep(1.5)
            else:
                print(f"Failed to fetch data for item {item_id} after {retries} attempts. Error: {str(e)}")
                return None
    return None

def get_spreadsheet_id(url_or_id: str) -> str:
    """
    Extract Spreadsheet ID from URL or return if already ID.
    """
    url_or_id = url_or_id.strip()
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url_or_id)
    if match:
        return match.group(1)
    return url_or_id

def authenticate_google(client_secret_path: str, token_path: str, port: int = 0) -> Credentials:
    """
    Authenticate with Google OAuth 2.0. Runs local server if user consent needed.
    """
    creds = None
    if os.path.exists(token_path):
        try:
            creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        except Exception:
            pass
            
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                creds = None
        
        if not creds:
            if not os.path.exists(client_secret_path):
                raise FileNotFoundError(
                    f"Không tìm thấy file JSON Credentials tại: {client_secret_path}\n"
                    "Vui lòng tải về từ Google Cloud Console và đặt vào thư mục config."
                )
            
            flow = InstalledAppFlow.from_client_secrets_file(client_secret_path, SCOPES)
            # Port=0 lets the system pick any available port automatically
            creds = flow.run_local_server(
                port=port, 
                prompt='consent', 
                authorization_prompt_message='Vui lòng cấp quyền trên trình duyệt web của bạn.'
            )
            
        # Save credentials for future use
        with open(token_path, 'w') as token_file:
            token_file.write(creds.to_json())
            
    return creds

def get_sheet_existing_ids(service, spreadsheet_id: str, sheet_name: str) -> Set[str]:
    """
    Fetch all Item IDs in column A from the specified sheet to build cache.
    """
    existing_ids = set()
    try:
        # Read column A
        range_name = f"'{sheet_name}'!A:A"
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, 
            range=range_name
        ).execute()
        
        rows = result.get('values', [])
        if rows:
            # Skip header row (index 0) and gather IDs
            for r in rows[1:]:
                if r:
                    existing_ids.add(str(r[0]).strip())
    except HttpError as e:
        # If the sheet doesn't exist, we will create it later, so return empty set
        if e.resp.status != 404:
            print(f"Error fetching existing IDs: {str(e)}")
    except Exception as e:
        print(f"Error reading sheet: {str(e)}")
        
    return existing_ids

def ensure_sheet_and_headers(service, spreadsheet_id: str, sheet_name: str) -> bool:
    """
    Checks if sheet exists. If not, creates it. Then ensures headers are present.
    """
    headers = [
        "Mã SP (Item ID)",
        "Mã Shop (Shop ID)",
        "Tên sản phẩm",
        "Giá (đ)",
        "Lượt bán",
        "Đường dẫn sản phẩm",
        "Hoa hồng (đ)",
        "Hoa hồng người bán (đ)",
        "Hoa hồng Shopee (đ)"
    ]
    
    try:
        # Get spreadsheet metadata
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        sheet_titles = [s.get('properties', {}).get('title') for s in sheets]
        
        # Create sheet if not exists
        if sheet_name not in sheet_titles:
            batch_update_request_body = {
                'requests': [{
                    'addSheet': {
                        'properties': {'title': sheet_name}
                    }
                }]
            }
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, 
                body=batch_update_request_body
            ).execute()
            
            # Write headers
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_name}'!A1",
                valueInputOption="RAW",
                body={"values": [headers]}
            ).execute()
            return True
            
        # If sheet exists, check if headers are set
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, 
            range=f"'{sheet_name}'!A1:I1"
        ).execute()
        
        rows = result.get('values', [])
        if not rows or not rows[0]:
            # Write headers if first row is empty
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_name}'!A1",
                valueInputOption="RAW",
                body={"values": [headers]}
            ).execute()
            
    except Exception as e:
        print(f"Error ensuring sheet/headers: {str(e)}")
        return False
    return True

def push_to_google_sheet(
    service, 
    spreadsheet_id: str, 
    sheet_name: str, 
    rows_to_append: List[Dict[str, Any]]
) -> int:
    """
    Appends list of dict products to Google Sheet.
    Returns the number of successfully pushed rows.
    """
    if not rows_to_append:
        return 0
        
    ensure_sheet_and_headers(service, spreadsheet_id, sheet_name)
    
    # Map dict key to correct column sequence
    keys_order = [
        "Mã SP (Item ID)",
        "Mã Shop (Shop ID)",
        "Tên sản phẩm",
        "Giá (đ)",
        "Lượt bán",
        "Đường dẫn sản phẩm",
        "Hoa hồng (đ)",
        "Hoa hồng người bán (đ)",
        "Hoa hồng Shopee (đ)"
    ]
    
    values = []
    for item in rows_to_append:
        row = [item.get(k, "") for k in keys_order]
        values.append(row)
        
    range_name = f"'{sheet_name}'!A:A"
    body = {
        'values': values
    }
    
    try:
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body=body
        ).execute()
        return len(values)
    except Exception as e:
        print(f"Error appending values to GSheets: {str(e)}")
        raise e

def export_to_excel(crawled_data: List[Dict[str, Any]], export_path: str):
    """
    Export collected data list to Excel spreadsheet file.
    """
    if not crawled_data:
        return
        
    df = pd.DataFrame(crawled_data)
    cols = [
        "Mã SP (Item ID)",
        "Mã Shop (Shop ID)",
        "Tên sản phẩm",
        "Giá (đ)",
        "Lượt bán",
        "Đường dẫn sản phẩm",
        "Hoa hồng (đ)",
        "Hoa hồng người bán (đ)",
        "Hoa hồng Shopee (đ)"
    ]
    
    df = df.reindex(columns=cols)
    df.to_excel(export_path, index=False)
