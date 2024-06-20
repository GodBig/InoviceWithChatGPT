import pytesseract
from PIL import Image
from langchain.agents import Agent, Tool
from langchain.prompts import ChatPromptTemplate
from langchain.chains import LLMChain, SimpleSequentialChain
from langchain.llms import OpenAI


# 设置Tesseract的路径（如果需要）
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# 定义OCR工具类
class OCRTool(Tool):
    def __init__(self, name, description, func):
        super().__init__(name, description)
        self.func = func

    def run(self, image_path):
        return self.func(image_path)


# 示例函数，使用Tesseract进行OCR
def ocr_invoice(image_path):
    img = Image.open(image_path)
    text = pytesseract.image_to_string(img)
    return text


# 创建OCR工具实例
ocr_tool = OCRTool(
    name="ocr_tool",
    description="A tool to extract text from invoice images.",
    func=ocr_invoice
)

# 加载OpenAI的语言模型
llm = OpenAI(api_key="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")

# 创建解析发票信息的链
parse_prompt = ChatPromptTemplate.from_template(
    "Firstly, I'm from China State Construction Middle East Company. I've received many invoices and work completion confirmation forms. I scanned them all together into images and used Tesseract OCR to recognize the content in the images. However, the image content is somewhat chaotic.Please assist me in identifying, correcting, and even filling in crucial information." +
    "Secondly, document type, purchasing department or project, invoice date, invoice number, LPO code, TRN code, and supplier name from the recognized content, and enter the results into the JSON format below." +
    str({
        "type": "",
        "project": "",
        "date": "",
        "invoice": "",
        "lpo": "",
        "trn": "",
        "supplier": "",
    }) +
    "If the provided content does not contain the required information to fill in the JSON format.Just give me the raw Json without any update."
    "This is the content from Tesseract OCR:"
)
parse_chain = LLMChain(llm=llm, prompt=parse_prompt)

# 将OCR工具和解析链组合成一个链式任务
invoice_chain = SimpleSequentialChain(chains=[ocr_tool, parse_chain])

# 创建一个Agent实例，并添加工具和链
agent = Agent(
    name="invoice_agent",
    description="An agent that can extract information from invoices.",
    tools=[ocr_tool],
    chains={"extract_invoice_info": invoice_chain}
)

# 使用Agent执行任务
image_path = "path_to_your_invoice_image.jpg"
query = {"image_path": image_path}
response = agent.run(query)
print(response)
