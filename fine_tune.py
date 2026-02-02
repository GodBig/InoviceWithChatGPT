
"""
Fine-tuning Pipeline Demo
-------------------------
This script demonstrates the process of:
1.  **Synthetic Data Generation**: Creating training data when real data is scarce.
2.  **Fine-tuning Preparation**: Formatting data for OpenAI/Llama training.
3.  **Model Training**: (Simulated) launching a training job.

Usage: This is a teaching module to show how domain-specific models are adapted.
"""

import json
import random
import time
import os

def generate_synthetic_invoice():
    """Generates a random invoice text and corresponding JSON label."""
    vendors = ["Alpha Corp", "Beta Industries", "Gamma Logistics", "Delta Services"]
    currencies = ["USD", "AED", "CNY", "EUR"]
    
    vendor = random.choice(vendors)
    inv_num = f"INV-{random.randint(1000, 9999)}"
    amount = round(random.uniform(100.0, 5000.0), 2)
    currency = random.choice(currencies)
    date = f"2024-{random.randint(1,12):02d}-{random.randint(1,28):02d}"
    
    # 1. Create the unstructured text representation (Input)
    raw_text = f"""
    INVOICE RECEIPT
    ----------------
    Issued By: {vendor}
    Date Of Issue: {date}
    Reference: {inv_num}
    
    Grand Total: {currency} {amount}
    
    Thank you for your business.
    """
    
    # 2. Create the structured label (Output)
    label_json = {
        "vendor": vendor,
        "date": date,
        "bg_number": inv_num,
        "amount": amount,
        "currency": currency
    }
    
    return raw_text, label_json

def create_training_dataset(num_samples=10):
    """Creates a JSONL file suitable for OpenAI fine-tuning."""
    print(f"Generating {num_samples} synthetic training samples...")
    training_data = []
    
    for _ in range(num_samples):
        text, label = generate_synthetic_invoice()
        
        # OpenAI Chat Format
        entry = {
            "messages": [
                {"role": "system", "content": "You are an invoice extraction assistant."},
                {"role": "user", "content": f"Extract data from this invoice:\n{text}"},
                {"role": "assistant", "content": json.dumps(label)}
            ]
        }
        training_data.append(entry)
        
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    output_dir = os.path.join(base_dir, "data", "training_data")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    output_path = os.path.join(output_dir, "fine_tune_job.jsonl")
    with open(output_path, "w", encoding="utf-8") as f:
        for item in training_data:
            f.write(json.dumps(item) + "\n")
            
    print(f"Training data saved to {output_path}")
    return output_path

def start_fine_tuning(file_path):
    """Simulates the API call to start training."""
    print("\n--- Initiating Fine-Tuning Job ---")
    print(f"Uploading file: {file_path}...")
    time.sleep(1)
    
    print("Job ID: ft-job-xyz123abc")
    print("Base Model: gpt-3.5-turbo")
    print("Epochs: 3")
    print("Status: Queued...")
    
    for i in range(3):
        time.sleep(1)
        print(f"Status: Training (Epoch {i+1}/3)... loss={random.uniform(0.1, 0.5):.4f}")
        
    print("Status: Succeeded")
    print("New Model ID: ft:gpt-3.5-turbo:my-org::model-v1")

if __name__ == "__main__":
    # 1. Generate Data
    dataset_path = create_training_dataset(num_samples=5)
    
    # 2. Train Model
    start_fine_tuning(dataset_path)
