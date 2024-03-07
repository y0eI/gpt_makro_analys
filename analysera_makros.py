import os
from glob import glob
from openai import OpenAI
from oletools.olevba import VBA_Parser

# OpenAI API nyckel
nyckel = OpenAI(api_key='')  


# Mapp där .doc filerna är sparade
mapp = ''

# Hitta alla .doc filer
doc_fil = os.path.join(mapp, '*.doc')
alla_doc_filer = glob(doc_fil)

# Hantera makros
for enskild_fil in alla_doc_filer:
    vba_parser = VBA_Parser(enskild_fil)

    alla_macros = ""
    # Se om det finns makros i file
    if vba_parser.detect_vba_macros():
        for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
            alla_macros += "Macro in: " + filename + "\n" + vba_code + "\n" + "-" * 40 + "\n"

    # Skicka makros till GPT
    if alla_macros:
        svar = nyckel.chat.completions.create(
            model="gpt-3.5-turbo-0125",
            messages=[
                {"role": "user", "content": "Analyze the following VBA code to determine its potential for harmful actions. Provide a percentage estimate of how likely it is to be malicious. Additionally, suggest structured actions that should be taken based on your analysis:\n\n" + alla_macros}
            ]
        )

        # Skapa output filen
        base_namn = os.path.basename(enskild_fil)
        output_fil_namn = f"svar_{base_namn}.txt"
        output_fil_path = os.path.join(mapp, output_fil_namn)

        # Skriv ut GPTs svar till .txt filen
        with open(output_fil_namn, 'w') as output_fil:
            output_fil.write("GPTs analys:\n")
            output_fil.write(svar.choices[0].message.content + "\n")


    vba_parser.close()
