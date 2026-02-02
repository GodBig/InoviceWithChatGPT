
"""
Intelligent Invoice Processing Agent
------------------------------------
This module implements a ReAct (Reasoning + Acting) agent to extract structured data 
from unstructured invoice documents.

Key Features:
1.  **Iterative Reasoning**: The agent plans its steps (e.g., "First read PDF, then validate tax ID").
2.  **Tool Use**: Access to OCR, Web Search (simulated), and Database validation.
3.  **Reflection**: Self-correction if the output JSON doesn't match schema.
"""

import json
from typing import Dict, Any, List
from datetime import datetime

# Mock classes to simulate LangChain/OpenAI behavior for demonstration
class MockLLM:
    def predict(self, prompt: str) -> str:
        """
        Simulates LLM response based on the prompt content.
        In a real scenario, this would call OpenAI/Azure/LocalLLM.
        """
        if "EXTRACT_JSON" in prompt:
            return json.dumps({
                "invoice_number": "INV-2024-001",
                "date": "2024-05-20",
                "total_amount": 1500.00,
                "currency": "USD",
                "vendor_name": "Global Tech Solutions",
                "tax_id": "TRN-882211"
            })
        elif "VALIDATE_TAX" in prompt:
            return "Tax ID format seems valid for UAE region."
        else:
            return "I will proceed to extract information from the document."

class InvoiceAgent:
    def __init__(self, llm_api_key: str = None):
        self.llm = MockLLM()
        self.history = []

    def log_process(self, step: str, details: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        print(f"[{timestamp}] [AGENT] {step}: {details}")
        self.history.append({"step": step, "time": timestamp, "details": details})

    def tool_ocr(self, file_path: str) -> str:
        """Simulates OCR Text Extraction"""
        self.log_process("TOOL_USE", f"Running OCR on {file_path}...")
        # In real code: pytesseract.image_to_string(Image.open(file_path))
        return """
        INVOICE
        Vendor: Global Tech Solutions
        Date: 20 May 2024
        Invoice #: INV-2024-001
        TRN: TRN-882211
        
        Description      Qty    Price    Total
        --------------------------------------
        Consulting Svc    1     1500     1500
        
        Total: $1,500.00
        """

    def tool_validate_tax_id(self, tax_id: str) -> bool:
        """Simulates Tax ID validation against an external API"""
        self.log_process("TOOL_USE", f"Validating Tax ID: {tax_id}...")
        return tax_id.startswith("TRN-")

    def run_workflow(self, file_path: str) -> Dict[str, Any]:
        """
        Executes the Agentic Workflow:
        1. Perception (OCR)
        2. Reasoning (Planning)
        3. Extraction
        4. Validation
        5. Reflection/Correction
        """
        print(f"--- Starting Agent Workflow for {file_path} ---")
        
        # Step 1: Perception
        raw_text = self.tool_ocr(file_path)
        
        # Step 2: Reasoning & Extraction
        self.log_process("THOUGHT", "Raw text obtained. Needs to be structured into JSON.")
        prompt = f"""
        Task: Extract invoice details into JSON.
        Fields required: invoice_number, date, total_amount, currency, vendor_name, tax_id.
        Raw Text:
        {raw_text}
        EXTRACT_JSON
        """
        response = self.llm.predict(prompt)
        
        try:
            data = json.loads(response)
            self.log_process("ACTION", f"Extracted Data: {data}")
        except json.JSONDecodeError:
            self.log_process("ERROR", "LLM returned invalid JSON. Retrying...")
            # Logic to retry with correction prompt would go here
            return {}

        # Step 3: Verification (Agent uses tools to check its own work)
        tax_id = data.get("tax_id")
        if tax_id:
            is_valid = self.tool_validate_tax_id(tax_id)
            if not is_valid:
                self.log_process("REFLECTION", "Tax ID appears invalid. Flagging for human review.")
                data["warnings"] = ["Invalid Tax ID format"]
            else:
                self.log_process("REFLECTION", "Tax ID verified successfully.")
        
        self.log_process("COMPLETE", "Workflow finished.")
        return data

if __name__ == "__main__":
    # Demonstration
    agent = InvoiceAgent(llm_api_key="sk-mock-key")
    result = agent.run_workflow("sample_invoice.pdf")
    print(f"\nFinal Result:\n{json.dumps(result, indent=2)}")
