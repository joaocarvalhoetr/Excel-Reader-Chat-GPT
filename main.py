import os
import dotenv
import pandas as pd
from openai import OpenAI
import logging
import json

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Load environment variables from .env file
dotenv.load_dotenv()

API_KEY = os.getenv('OPENAI_API_KEY')
if not API_KEY:
    logging.error("API key not found. Please set the OPENAI_API_KEY environment variable.")
    exit(1)

# Read the EXCEL file
try:
    df = pd.read_excel('data.xlsx')
    logging.info("Excel file read successfully.")
except Exception as e:
    logging.error(f"Failed to read Excel file: {e}")
    exit(1)

# Store each excel line in a list
raw_data = df.values.tolist()
logging.info(f"Raw data extracted: {raw_data}")  # Log the entire data

# Create an OpenAI client
client = OpenAI(api_key=API_KEY)

# Prepare a list to hold the processed results
results = []

# Iterate over each row in raw_data
for index, row in enumerate(raw_data):
    # Prepare the message content
    message_content = str(row[1] ) + str(row[3]) + str(row[7])  # Ensure it's a string

    # Create the prompt for the OpenAI GPT model
    prompt = f"""
    You are a helpful assistant that will convert a list into info in the following format: 
    Empresa, Nome, Ultimo Nome, Posição, Email, Telefone. 
    Please output this in JSON format nothing else.
    Here is the list: {message_content}
    """

    # Make the API call to OpenAI using the new interface
    try:
        chat_completion = client.chat.completions.create(
            messages=[
                {"role": "user", "content": prompt}
            ],
            model="gpt-4o-mini",  # Use "gpt-4" if you have access
        )
        logging.info(f"API call successful for row {index}.")
    except Exception as e:
        logging.error(f"API call failed for row {index}: {e}")
        continue  # Skip to the next row if the API call fails

    # Extract the assistant's response
    assistant_response = chat_completion.choices[0].message.content.strip()

    # Log and print the response from OpenAI
    logging.info(f"Response from OpenAI for row {index}: {assistant_response}")

    # Convert the response to a dictionary
    try:
        # Make sure to clean up the JSON string
        cleaned_response = assistant_response.replace("```json", "").replace("```", "").strip()
        data_dict = json.loads(cleaned_response)
        logging.info("Response converted to dictionary.")
        results.append(data_dict)  # Append the result to the list
    except json.JSONDecodeError as e:
        logging.error(f"Failed to convert response to dictionary for row {index}: {e}")
        logging.error(f"Response content: {assistant_response}")  # Log the problematic response
        continue  # Skip to the next row if conversion fails

# Convert results list to DataFrame
try:
    result_df = pd.DataFrame(results)
    logging.info("All responses converted to DataFrame.")
except Exception as e:
    logging.error(f"Failed to convert results to DataFrame: {e}")
    exit(1)

# Save DataFrame to CSV
try:
    result_df.to_csv('output.csv', index=False)
    logging.info("Data saved to output.csv successfully.")
except Exception as e:
    logging.error(f"Failed to save data to CSV: {e}")
    exit(1)
