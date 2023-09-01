# pptx-translator

Python script that translates pptx files using Amazon Translate service. This script differs from the [original]([url](https://github.com/aws-samples/pptx-translator)) in a few ways. The major difference is that the translated text appends, rather than replaces, the original text. Another difference is on the input and output file management. Instead of requesting the file path, the script automatically works on the ppt files in the workspace directory, and saves the output files in the output sub-directory of the workspace directory.
![image](https://github.com/wenchaoliu-93/pptx-translator-bilingual/assets/121582343/afeb667d-4d71-4630-8550-82d42799af21)
Video: http://www.youtube.com/watch?v=hHuFEh-w1dE

## Installation

```
$ virtualenv venv
$ source venv/bin/activate
$ pip install -r requirements.txt
```

## Usage
```
$ python pptx-translator.py --help
usage: Translates pptx files from source language to target language using Amazon Translate service
       [-h] [--terminology TERMINOLOGY]
       source_language_code target_language_code input_file_path

positional arguments:
  source_language_code  The language code for the language of the source text.
                        Example: en
  target_language_code  The language code requested for the language of the
                        target text. Example: pt

optional arguments:
  -h, --help            show this help message and exit
  --terminology TERMINOLOGY
                        The path of the terminology CSV file
```

## What is the program based off? 

The program is modified from an Amazon Web Services(AWS) sample that translates pptx files using Amazon Translate service. Here is documentation for Amazon Translate, which includes developer guide and API reference: Amazon Translate Documentation. 
 

## Execution Requirements 

There are three major pieces of requirements to execute the program. First, Python 3 installation, as the program is written in Python. Second, python-pptx library. It is the program backbone for manipulating PowerPoint files. Here is the library documentation: python-pptx — python-pptx 0.6.21 documentation. Three, a pair of AWS access key ID and secret access key. The pair of keys needs to be passed into the code. 

 

## Program use case 

The use case is to translate pptx slides that contain only one language into dual-language slides. Specifically, the program translates English text into Chinese, and append the translated text at the end of each paragraph. With slight modifications, however, the pair of languages can be of any combination, so long as they are supported by AWS.  

 

## Instructions 

Before execution, the source pptx files should be placed in the workspace folder. The one argument is optional, which is path of the terminology CSV file.  Once executed, the program will output one pptx file for each source file. 

Here is documentation on translation customizations: Customizing your translations with Amazon Translate - Amazon Translate. There are five customization settings: do-not-translate tags, custom terminology, profanity, formality, parallel data. 

## Contact

wenchaoliu93@gmail.com

## Features 

Two major features are worth noting. First, in many instances, appending text inevitably leads to overflow. The program automatically resizes the text to fit the text frame. This feature does need occasional tweaks. Second, the program can skip certain text for translation. It comes in handy in situations where there is reoccurring text that needs no translation, such as text that displays author or institution information. 
