# PowergenX

This is a project utilizes GPT-4 to revolutionize PowerPoint presentation creation. It would generate powerpoint automatically according to user's preferences and proposals.

## Description

Our product would begin with user input, including a proposal document and specific presentation requirements like format and length. This input feeds into the GPT-4 Module, which processes the information and generates content that is then formatted for PowerPoint creation. Followed by a PowerPoint Generator Module which uses this formatted output to craft a presentation with various slide layouts. The user reviews the generated presentation and can request changes, leading to iterative refinements. The process loops until the user is satisfied, culminating in the final presentation ready for use.
## Getting Started

### Dependencies

* WindowsOs
Python library openai, python-docx, python-pptx, IPython needed. Or you could directly install them in Jupyter notebook ```PowerGenXdemo.ipynb```

### What'd you need  
Our product would need at least one document or lines of description upon the presentation you are seeking for. For example, you could upload your ppt requirements and your proposal document, then our model would help you generate a related powerpoint for you.

### Executing program
* How to run the program:
    * Download the zip file through web or use ```git clone https://github.com/ece1786-2023/PowerGenX.git```
    * Open the Jupyter Noterbook ```PowerGenXdemo.ipynb``` with preferred inerpreter
    * Follow the instruction within the notebook to run cells  
    * Activate the Gradio frontend and upload your files according to the frontend  
    * Generate the ppt by clicking the button, and download it in the output window, remember to change the ppt name!  

## Help

If you are unable to open the Jupyter noterbook on local machine, you could use the google Colab [(Colab Research)](https://colab.research.google.com/) to handle the notebook. You need to upload the notebook on colab and open it in web, then follow the instructions within the notebook to run cells and activate the Gradio frontend.

## Authors


Yuquan (Johnny) Gan [@yuquan](https://github.com/supremejohnny)   
Tianze (Alex) Wang

## Version History

* 0.1
    * Initial Release, with basic gradio frontend and functionality


## Acknowledgments

-
