# -*- coding: utf-8 -*-
"""

"""

import os
import sys
import argparse
from pyexcel_xlsx import get_data

def main(input, media, output, category, system):

    if not os.path.exists(output):
        os.makedirs(os.path.join(output, "media"))

    #media: %APPDATA%\Anki2\{profile}\collection.media
    #fields: pytanie, odpowiedz, punkty
    template = "{0}\t{1}\t{2}\n"
    outfile = open(os.path.join(output, "prawo-jazdy-pytania.txt"), "w", encoding="utf-8")
    copy_command = 'copy' if system == 'windows' else 'cp'

    data = get_data(input)
    for item in data["pytania"]:
        kategorie = item[17].split(',')
        if category in kategorie:
            #numer_pytania = item[0]
            #tresc_pytania = item[1]
            pytanie = item[1]
            odpowiedz_a = item[2]
            odpowiedz_b = item[3]
            odpowiedz_c = item[4]
            odpowiedz_poprawna = item[13]
            media_file = item[14]
            liczba_punktow = item[16]
            
            if media_file:
                    
                output_media_file = os.path.join(output, "media", media_file)
                
                if not os.path.exists(output_media_file):
                    input_media_file = os.path.join(os.path.dirname(input), "media", media_file)
                    if not os.path.exists(input_media_file):
                        raise Exception("Missing media file: " + input_media_file)
                    
                    os.system(copy_command + ' "{0}" "{1}"'.format(input_media_file, output_media_file))
                    if not os.path.exists(output_media_file):
                        raise Exception("Cannot copy media file: " + input_media_file)
            question = '<div class="pytanie">' + pytanie + '</div>'
            if media_file:
                question += '<div class="media">'
                if media_file.endswith(('.mp4','.wmv','.avi')):
                    question += '[sound:' + media_file + ']'
                elif media_file.endswith('.jpg'):
                    question += '<img src="' + media_file + '" />'
                else:
                    raise Exception("Unexpected extension in media file: " + media_file)
                question +='</div>'
            if odpowiedz_a or odpowiedz_b or odpowiedz_c:
                question += '<div class="odpowiedzi">'
                if odpowiedz_a:
                    question += '<span class="wariant">A</span>: ' + odpowiedz_a + '<br>'
                if odpowiedz_b:
                    question += '<span class="wariant">B</span>: ' + odpowiedz_b + '<br>'
                if odpowiedz_c:
                    question += '<span class="wariant">C</span>: ' + odpowiedz_c + '<br>'
                question +='</div>'
            
            answer = '<div class="odpowiedz">'
            if odpowiedz_poprawna == 'T':
                answer += 'TAK'
            elif odpowiedz_poprawna == 'N':
                answer += 'NIE'
            elif odpowiedz_poprawna == 'A':
                answer += 'A: ' + odpowiedz_a
            elif odpowiedz_poprawna == 'B':
                answer += 'B: ' + odpowiedz_b
            elif odpowiedz_poprawna == 'C':
                answer += 'C: ' + odpowiedz_c
            else:
                raise Exception("Unexpected answer: " + odpowiedz_poprawna)
            answer += '</div>'
            outfile.write(template.format(question, answer, liczba_punktow))
    #print(media)
    #print(output)
    outfile.close()
    return 0

def parse_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input', 
        help='Input XLSX file name', required=True)
    parser.add_argument('-m', '--media', 
        help='Media directory', required=True)
    parser.add_argument('-o', '--output', 
        help='Output directory', required=True)
    parser.add_argument('-c', '--category', 
        help='License category', required=True)
    parser.add_argument('-s', '--system', 
        help='Operating system', required=False, default='windows', choices=['windows','linux'])
    args = parser.parse_args()
    return args

if __name__ == '__main__':
    args = parse_arguments()
    try:
        sys.exit(main(input = args.input, media = args.media, output = args.output, category = args.category, system = args.system))
    except Exception as exc:
        print(exc, file=sys.stderr, end="\n")
        sys.exit(1)
    
