import os
import subprocess
import textwrap
from datetime import datetime
from openai import OpenAI
from lxml import etree
from fpdf import FPDF
from PyPDF2 import PdfMerger

# Insert OpenAI API key here
client = OpenAI(api_key = "API_KEY_HERE")

# Remove namespace prefixes
def remove_namespaces(element):
    for elem in element.getiterator():
        if not (isinstance(elem, etree._Comment) or isinstance(elem, etree._ProcessingInstruction)):
            elem.tag = etree.QName(elem).localname
    etree.cleanup_namespaces(element)

# Extract policy information from a given XML file
def extract_policies(input_file_path, formatted_file_path, compressed_file_path):
    tree = etree.parse(input_file_path)
    root = tree.getroot()

    new_root = etree.Element("GPO")
    #xmlns="http://www.microsoft.com/GroupPolicy/Settings",
    #xmlns_ns0="http://www.microsoft.com/GroupPolicy/Settings/Security",
    #xmlns_ns1="http://www.microsoft.com/GroupPolicy/Settings/Registry",
    #xmlns_xsi="http://www.w3.org/2001/XMLSchema-instance")

    # Find the Computer Configuration and User Configuration elements, regardless of namespace
    computer = root.find(".//{*}Computer")
    user = root.find(".//{*}User")
    if computer is None or user is None:
        return False

    # Remove <Explain> sections from the Computer and User elements
    def remove_explain_sections(element):
        if element is not None:
            explains = element.findall(".//{*}Explain")
            for explain in explains:
                explain.getparent().remove(explain)

    remove_explain_sections(computer)
    remove_explain_sections(user)

    # Append the Computer and User data to the new root element
    if computer is not None:
        new_root.append(computer)
    if user is not None:
        new_root.append(user)

    # Remove namespaces from all tags in the tree
    remove_namespaces(new_root)

    # Formats the XML file with regular indents for logging purposes
    def pretty_write(element, file):
        def recursive_write(elem, file, level=0):
            indent = "  " * level
            file.write(f"{indent}<{elem.tag}>")
            if elem.text and elem.text.strip():
                file.write(f"{elem.text.strip()}")
            file.write("\n")
            for child in elem:
                recursive_write(child, file, level + 1)
            file.write(f"{indent}</{elem.tag}>\n")

        recursive_write(element, file)

    # Formats the XML file to one line with no whitespace, reducing the amount of characters sent to OpenAI's API
    def compress_write(element, file):
        def recursive_write(elem, file):
            file.write(f"<{elem.tag}>")
            if elem.text and elem.text.strip():
                file.write(f"{elem.text.strip()}")
            file.write("")
            for child in elem:
                recursive_write(child, file)
            file.write(f"</{elem.tag}>")

        recursive_write(element, file)

    # Write the formatted XML to the output file
    with open(formatted_file_path, "w", encoding="utf-8") as formatted_file:
        formatted_file.write("<?xml version='1.0' encoding='utf-8'?>\n")
        pretty_write(new_root, formatted_file)
        formatted_file.close()

    # Write the compressed XML to a separate output file
    with open(compressed_file_path, "w", encoding="utf-8") as compressed_file:
        compressed_file.write("<?xml version='1.0' encoding='utf-8'?>")
        compress_write(new_root, compressed_file)
        compressed_file.close()

    print(f"Formatted XML file has been written to: {formatted_file_path}")
    print(f"Compressed XML file has been written to: {compressed_file_path}")

# Send the compressed XML file to OpenAI and ask it for feedback
def query_openai(file_path):
    try:
        with open(file_path, "r") as file:
            xml_content = file.read()
    except FileNotFoundError:
        # This can only happen if the user removes the compressed XML from the directory before sending it to the AI
        print("Compressed file could not be found")
        return
    
    response = client.chat.completions.create(
        model = "gpt-4o",
        messages = [
            {"role": "system", "content": "You are an assistant that provides feedback and suggestions on group policy objects."},
            {"role": "user", "content": f"This is information taken from an XML GPO report, please give feedback and suggestions on the group policy object based on best practices:\n\n{xml_content}"}
        ]
    )

    response_text = response.choices[0].message.content.encode("latin-1", "replace").decode("latin-1")
    return response_text

# Write AI response to log file
def write_log(text):
    file = open(log_file, "w")
    now = datetime.now()
    formatted_now = now.strftime("%Y-%m-%d %H:%M:%S")
    file.write(f"Date: {formatted_now}\nReport file: {input_file}\n\n")
    file.write(f"{text}\n\n")
    file.close()

# Use log_file to create a new PDF
def generate_pdf():
    # Setup PDF settings and log_file
    pt_to_mm = 0.4
    fontsize_pt = 10
    fontsize_mm = fontsize_pt * pt_to_mm
    character_width_mm = 7 * pt_to_mm
    width_text = 325 / character_width_mm
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(True, margin=10)
    pdf.add_page()
    pdf.set_font(family="Arial", size=fontsize_pt)

    file = open(log_file)
    text = file.read()
    file.close()
    splitted = text.split("\n")
    for line in splitted:
        lines = textwrap.wrap(line, width_text)
        if len(lines) == 0:
            pdf.ln()
        for wrap in lines:
            pdf.cell(0, fontsize_mm, wrap, ln=1)

    if os.path.isfile(output_file):
        # PDF log exists, create a throwaway one then merge
        throwaway = "Throwaway.pdf"
        pdf.output(throwaway, "F")
        merge = PdfMerger()
        merge.append(output_file)
        merge.append(throwaway)
        merge.write(output_file)
        merge.close()
        os.remove(throwaway)
    else:
        # PDF log does not exist yet, create new one with proper name
        pdf.output(output_file, "F")
    print(f"\nOutput written to {output_file}")

# Use powershell to generate a GPO report
def generate_report():
    user_input = input("Enter the name of an existing group policy object:\n")
    exception = "if($?){echo Success}else{echo Error}"
    subprocess.call(f"powershell.exe Get-GPOReport -Name '{user_input}' -ReportType XML -Path '{user_input} Report.xml'; {exception}")
    print("\n***Returning to menu***")

# Take in an XML GPO report and format/compress it
def input_report():
    global input_file
    while True:
        user_input = input("Enter path to an XML GPO Report:\n")
        if not os.path.exists(user_input):
            print("Please enter a valid path")
        elif not user_input.lower().endswith(".xml"):
            print("Please enter a valid XML file")
        elif extract_policies(user_input, formatted_file, compressed_file) == False:
            print("Please enter a valid GPO Report")
        else:
            # Global input_file is to be used later in the txt and PDF log
            input_file = os.path.abspath(user_input)
            submit_report()
            break

# Send a compressed report to OpenAI API
def submit_report():
    while True:
        user_prompt = input("Send compressed report to AI for analysis? (Y/N): ").lower()
        if user_prompt == "yes" or user_prompt == "y":
            print("Waiting for AI response...")
            response = query_openai(compressed_file)
            print(response)
            write_log(response)
            generate_pdf()
            print("\n***Returning to menu***")
            break
        elif user_prompt == "no" or user_prompt == "n":
            print("\n***Returning to menu***")
            break
        else:
            print("Please enter yes or no")

def menu_loop():
    while True:
        print("\n1. Generate a GPO report\n2. Submit a GPO report for analysis\n3. Close program")
        menu_input = input("Please select an option (1-3): ")
        if menu_input == "1":
            generate_report()
        elif menu_input == "2":
            input_report()
        elif menu_input == "3":
            print("Closing program...")
            exit()
        else:
            print("Invalid input, please enter a number (1-3)")

# Formatted XML file path, for demonstration purposes
formatted_file = "./Formatted Report.xml"

# Compressed XML file path
compressed_file = "./Compressed Report.xml"

# Log file path of last response
log_file = "./Text Log.txt"

# PDF log of all responses
output_file = "./Output Log.pdf"


menu_loop()
