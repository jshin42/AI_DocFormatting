DocFormatting with OpenAI GPT

This Python script, docformatting.py, leverages OpenAI's GPT model to enhance the readability and conciseness of text within Microsoft Word documents. It reads a specified Word document, processes its content through the GPT model to clean up and refine the text, and then saves the enhanced text back into a Word document. This tool is ideal for improving document quality, making complex texts more accessible, and tailoring content towards specific audiences.
Features

    Text Enhancement: Utilizes OpenAI's GPT-4 model to rewrite and enhance text, making it more readable and concise.
    Document Processing: Works with Microsoft Word (.docx) documents, reading and saving content seamlessly.
    Batch Processing: Splits large documents into manageable chunks for processing, then recombines them, ensuring large files are handled efficiently.
    Customizable Prompts: Allows for the customization of the AI's prompt to tailor the text enhancement towards specific audiences or requirements.

Requirements

Before running this script, ensure you have the following prerequisites installed:

    Python 3.x
    docx library for handling Word documents.
    openai library for interacting with OpenAI's API.

These dependencies can be installed using pip:

bash

pip install python-docx openai

Additionally, you will need an OpenAI API key to use the GPT model. This key should be set in the API_KEY variable at the top of the script:

python

API_KEY = 'YOUR_OPENAI_API_KEY'

Usage

To use the script:

    Prepare Your Document: Place the Word document you wish to enhance in an accessible location and note its path.

    Configure the Script: Open doc_enhancer.py and set the INPUT_FILENAME, OUTPUT_FOLDER, and COMBINED_OUTPUT_FILENAME variables to match your document path and desired output locations.

    Run the Script: Execute the script from your terminal or command line:

    bash

    python doc_enhancer.py

    Review the Enhanced Document: Once processing is complete, open the combined output document to review the enhanced text.

Customization

You can customize how the script processes text by modifying the prompt in the enhance_text function. This allows you to target the text enhancement to specific audiences or to focus on particular aspects of readability and conciseness.
Limitations

    The script currently processes text only; images and other non-text elements in the document are not modified.
    The effectiveness of the text enhancement depends on the quality of the input text and the specific instructions given to the GPT model.

Contributing

Contributions are welcome! If you have suggestions for improving the script or extending its capabilities, please feel free to fork the repository, make your changes, and submit a pull request.
License

This script is provided under an open-source license. Please review the LICENSE file for more information.
