# pptx-translator

Python script that translates pptx files using Amazon Translate service. This script is built upon the [original AWS example](https://github.com/aws-samples/pptx-translator) and differs in a few ways. The major difference is that the translated text appends, rather than replaces, the original text. Another difference is on the input and output file management. Instead of requesting the file path, the script automatically works on the ppt files in the workspace directory, and saves the output files in the output sub-directory of the workspace directory.

![image](https://github.com/wenchaoliu-93/pptx-translator-bilingual/assets/121582343/afeb667d-4d71-4630-8550-82d42799af21)
[**Demo Video**](https://www.youtube.com/watch?v=hHuFEh-w1dE)

## Installation

```
$ virtualenv venv
$ source venv/bin/activate
$ pip install -r requirements.txt
```

## Execution Requirements 

There are three major pieces of requirements to execute the program. First, Python 3 installation, as the program is written in Python. Second, python-pptx library. It is the program backbone for manipulating PowerPoint files. Here is the library documentation: python-pptx — python-pptx 0.6.21 documentation. Last, a pair of AWS access key ID and secret access key. The pair of keys needs to be passed in the code. 

## Program use case 

The use case is to translate pptx slides that contain only one language into dual-language slides. Specifically, the program translates English text into Chinese, and append the translated text at the end of each paragraph. With slight modifications, however, the pair of languages can be of any combination, so long as they are supported by AWS.  

## Instructions 

Before execution, the source pptx files should be placed in the workspace folder. The one argument is optional, which is path of the terminology CSV file.  Once executed, the program will output one pptx file for each source file. 

## Features 

Two major features are worth noting. First, in many instances, appending text inevitably leads to overflow. The program automatically resizes the text to fit the text frame. This feature does need occasional tweaks; simply drag the textbox around would work. Second, the program can skip certain text for translation. It comes in handy in situations where there is reoccurring text that needs no translation, such as text that displays author or institution information. 

## Contact

wenchaoliu93@gmail.com
