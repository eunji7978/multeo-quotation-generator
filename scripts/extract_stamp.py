from PIL import Image
import numpy as np

def extract_red_stamp(image_path, output_path):
    print(f"Opening {image_path}...")
    img = Image.open(image_path).convert("RGBA")
    data = np.array(img)
    
    # Define red range
    # Stamps are usually reddish.
    # R > 100, G < 100, B < 100 is a rough approximation for deep red
    # Let's use a more flexible threshold.
    r, g, b, a = data.T
    
    # Condition: Red is dominant
    red_areas = (r > 130) & (g < 100) & (b < 100)
    
    # Create mask
    mask = np.zeros_like(r, dtype=bool)
    mask[red_areas] = True
    
    if not np.any(mask):
        print("No red stamp found.")
        return

    # Find bounding box
    coords = np.argwhere(mask.T) # Transpose back to match (y, x)
    # coords is (N, 2), where cols are y and x (or x and y depending on usage)
    # PIL/numpy coordinate mapping:
    # np.array shape is (Height, Width, 4)
    # r,g,b,a = data.T shape is (Width, Height). Wait, T swaps axes.
    # standard numpy: data[y, x]
    
    # easier way:
    rows = np.any(red_areas, axis=1) # Check rows that have red
    cols = np.any(red_areas, axis=0) # Check cols that have red
    
    y_min, y_max = np.where(rows)[0][[0, -1]]
    x_min, x_max = np.where(cols)[0][[0, -1]]
    
    print(f"Cropping to: x={x_min}~{x_max}, y={y_min}~{y_max}")
    
    # Crop
    stamp_crop = img.crop((x_min, y_min, x_max+1, y_max+1))
    
    # Optional: Make non-red pixels transparent? 
    # The user might want the white background or transparent. transparent is safer.
    # Let's process the cropped image to make light pixels transparent
    
    datas = stamp_crop.getdata()
    new_data = []
    for item in datas:
        # item is (r,g,b,a)
        # If it's white-ish, make transparent
        if item[0] > 200 and item[1] > 200 and item[2] > 200:
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)
            
    stamp_crop.putdata(new_data)
    
    stamp_crop.save(output_path)
    print(f"Saved stamp to {output_path}")

if __name__ == "__main__":
    # Path provided in conversation
    input_path = "/Users/kim-eunji/.gemini/antigravity/brain/bc417efa-8056-4313-9de0-4266cebf7a4c/uploaded_image_1_1767451018187.png"
    output_path = "assets/stamp.png"
    extract_red_stamp(input_path, output_path)
