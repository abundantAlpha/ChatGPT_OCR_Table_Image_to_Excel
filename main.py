import os
import base64
import requests
import pandas as pd
import json
from PIL import ImageGrab, Image
from io import BytesIO
from tkinter import Tk, messagebox


def convert_table_image_to_excel():

    def show_error_popup(message):
        root = Tk()
        root.withdraw()  # Hide the main Tkinter window
        messagebox.showerror("Error", message)
        root.destroy()

    def get_image_from_clipboard():
        try:
            img = ImageGrab.grabclipboard()
            if img is None:
                raise ValueError("No image found in clipboard!")
            if isinstance(img, Image.Image):
                if img.mode != "RGB":
                    img = img.convert("RGB")
                return img
            else:
                raise ValueError("Clipboard content is not an image!")
        except Exception as e:
            show_error_popup(f"Error retrieving image from clipboard: {e}")
            return None

    def create_table(image, api_key):
        try:
            buffer = BytesIO()
            image.save(buffer, format="JPEG")  # Save image in JPEG format
            image_bytes = buffer.getvalue()
            base64_image = base64.b64encode(image_bytes).decode('utf-8')

            system = [{"role": "system", "content": """
            You are an AI assistant with computer vision.
            You only output lists of dictionaries. Any other unit of communication is not allowed.
    
            Your built-in vision capabilities include:
            - extract numbers from image
            - understand table structures and logic
            - logical problem solving requiring reasoning and contextual consideration
            """.strip()
            }]

            user = [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "Return the contents of this image as lists of dictionaries with no formatting or special characters",
                        },
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/png;base64,{base64_image}"},
                        },
                    ],
                }
            ]

            params = {
                "model": "gpt-4o",
                "messages": system + user,
                "max_tokens": 1500, "top_p": 0.5, "temperature": 0.5,
            }

            headers = {
               "Content-Type": "application/json",
               "Authorization": f"Bearer {api_key}"
            }

            response = requests.post("https://api.openai.com/v1/chat/completions",
                                     headers=headers,
                                     json=params
                                     )

            if response.status_code != 200:
                raise ValueError(f"HTTP error {response.status_code}: {response.text}")

            response = response.json()['choices'][0]['message']['content']
            json_data = json.loads(response[7:len(response)-4])
            df = pd.DataFrame(json_data)
            return df

        except Exception as e:
            show_error_popup(f"Error creating table: {e}")
            return None

    # get user-defined variables
    with open("config.json", "r") as config_file:
        config = json.load(config_file)
        api_key = config["api_key"]
        output_file_path = config["output_file_path"]
        config_file.close()

    # create and open the CSV file
    try:
        image = get_image_from_clipboard()
        if image:
            my_df = create_table(image, api_key)
            if my_df is not None:
                my_df.to_csv(output_file_path)
                os.startfile(output_file_path)
            else:
                raise ValueError("Failed to create the DataFrame from the image.")
        else:
            raise ValueError("No valid image found to process.")
    except Exception as e:
        show_error_popup(f"Unexpected error: {e}")


convert_table_image_to_excel()
