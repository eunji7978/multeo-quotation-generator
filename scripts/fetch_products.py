import requests
from bs4 import BeautifulSoup
import json
import re
import os

def fetch_products():
    url = "https://multeo.kr/category/all/23/"
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36"
    }
    
    print(f"Fetching {url}...")
    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        print(f"Failed to fetch page: {response.status_code}")
        return []

    soup = BeautifulSoup(response.text, 'html.parser')
    
    products = []
    
    # Cafe24 usually uses .prdList or .xans-product-listnormal
    # Based on the text output we saw: "상품명 : ...", "판매가 : KRW ..."
    # Let's try to find the product items.
    
    # Strategy: Find all elements that might contain the product info.
    # Common classes in Cafe24: .item, .box, .description
    
    items = soup.select('.item') or soup.select('li > .box') or soup.find_all('div', class_='description')
    
    print(f"Found {len(items)} potential item blocks.")
    
    for item in items:
        # Debugging: Print item text to understand structure
        # print("DEBUG Item text:", item.get_text()) 

        name_tag = item.find('p', class_='name') or item.find('strong', class_='name')
        price_tag = item.find('li', class_='price') or item.find('span', class_='price')

        name = "Unknown"
        price = 0
        
        # Extract Name
        if name_tag:
            raw_name = name_tag.get_text(strip=True)
        else:
            raw_name = item.get_text(strip=True)

        # Look for "상품명 : Actual Name"
        # Regex: 상품명\s*[:]\s*(?P<name>[^\[\]\n]+)
        # Note: The chunk view showed format like "[상품명 : 미스트 찻잔]"
        name_match = re.search(r'상품명\s*[:]\s*([^\[\]\n\r]+)', raw_name)
        if name_match:
            name = name_match.group(1).strip()
        elif name_tag:
            # If explicit name tag exists but regex failed, trust the tag content (minus "상품명 :")
            name = raw_name.replace("상품명 :", "").replace("상품명", "").strip()
            
        # Cleanup: Remove "월 원" or "판매가" or "KRW" if they accidentally got into the name
        # This happens if we fall back to raw_name which is the whole text
        for marker in ["월 원", "판매가", "KRW", "상품명"]:
            if marker in name:
                name = name.split(marker)[0].strip()
        
        # Remove trailing junk often found in these raw texts
        name = re.sub(r'\s*[:]\s*$', '', name) # Remove trailing colon
            
        # Extract Price
        if price_tag:
             # Try to extract price
            price_text = ""
            if hasattr(price_tag, 'get_text'):
                price_text = price_tag.get_text(strip=True)
            else:
                price_text = str(price_tag).strip()
        else:
             price_text = item.get_text(strip=True)
        
        # Regex to find numbers after KRW or just numbers
        # Look for "KRW 50,000" or similar
        price_match = re.search(r'KRW\s*([\d,]+)', price_text)
        if not price_match:
             price_match = re.search(r'판매가\s*[:]\s*([\d,]+)', price_text)

        if price_match:
            price_str = price_match.group(1).replace(',', '')
            price = int(price_str)

        if name != "Unknown" and price > 0 and name != "상품명":
            # Remove any trailing "BEST" or other distinct markers if attached
            products.append({
                "name": name,
                "price": price,
                "category": "Unknown" # Placeholder
            })
            print(f"Found: {name} - {price}")

    return products

if __name__ == "__main__":
    products = fetch_products()
    
    output_path = "assets/products.json"
    os.makedirs("assets", exist_ok=True)
    
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(products, f, ensure_ascii=False, indent=2)
    
    print(f"Saved {len(products)} products to {output_path}")
