from Powergen import *
from generator import *
import re
import gradio as gr

def run_PowerGenX_code(code_block):
    # Use a regular expression to find the code starting with 'a = Poweregen()'
    # and remove Markdown backticks
    code_to_run = re.sub(r'^```python\n', '', code_block, flags=re.MULTILINE)
    code_to_run = re.sub(r'```$', '', code_to_run, flags=re.MULTILINE)

    # Make sure we start executing from the instantiation of Poweregen
    match = re.search(r'a = Poweregen\(\)', code_to_run)
    if match:
        code_to_run = code_to_run[match.start():]  # Extract from 'a = Poweregen()' to the end
        try:
            # Execute the extracted code
            exec(code_to_run, globals())
            print("Code executed successfully.")
        except Exception as e:
            print(f"Error executing the code: {e}")
    else:
        print("Could not find the starting point of Poweregen instantiation.")

#file_path1 = 'C:\\Users\\user\\Desktop\\研究生\\研一上\\1786h\\project\\PowergenX\\ExampleProposal.docx'
#proposal_text = read_word_document(file_path1)

def gradio_interface(proposal_file, num_slides, text_style, length, theme, layout):
  # Get the file extension
  filename = proposal_file.name
  file_extension = os.path.splitext(filename)[1].lower()

  # Read the file based on its type
  if file_extension == '.docx':
      proposal_text = read_word_document(proposal_file)
  elif file_extension == '.pdf':
      proposal_text = read_pdf_document(proposal_file)
  else:
      raise ValueError("Unsupported file type. Please upload a .docx or .pdf file.")
  presentation_code = generate_presentation_content(proposal_text, num_slides, text_style, length, theme, layout)
  run_PowerGenX_code(presentation_code)
  down = a.save_to_file()
  return down

iface = gr.Interface(
    fn=gradio_interface,
    inputs=[gr.File(label='Upload your file'),
        gr.Textbox(label='Number of slides'),
        gr.Radio(['concise', 'medium', 'detailed'], label='Text Style'),
        gr.Textbox(label='Length of presentation (in minutes)'),
        gr.Radio(['Default', 'Modern', 'Classic', 'Colorful'], label='PowerPoint Theme'),
        gr.Radio(['Standard', 'Creative', 'Professional', 'Minimalistic'], label='PowerPoint Layout')],
    #outputs=gr.Textbox(label="Generated PowerPoint, you need to change the file name by your self"),
    outputs=gr.File(label="Generated PowerPoint, you need to change the file name by your self"),
    title="PowergenX"
)
iface.launch()
