# -*- coding: utf-8 -*-
"""

"""

import os
import sys
import argparse
from pyexcel_xlsx import get_data

def main(input, media, output, category):

    if not os.path.exists(output):
        os.makedirs(os.path.join(output, "media"))

    #media: %APPDATA%\Anki2\{profile}\collection.media
    #fields: pytanie, odpowiedz, blok, punkty, komentarz, zrodlo
    template = "{0}\t{1}\t{2}\t{3}\t{4}\t{5}\n"
    outfile = open(os.path.join(output, "prawo-jazdy-pytania.txt"), "w", encoding="utf-8")
    
    data = get_data(input)
    for item in data["Treść pytania"]:
        kategorie = item[18].split(',')
        if category in kategorie:
            #nazwa_pytania = item[0]
            #numer_pytania = item[1]
            pytanie = item[2]
            odpowiedz_a = item[3]
            odpowiedz_b = item[4]
            odpowiedz_c = item[5]
            odpowiedz_poprawna = item[14]
            media_file = item[15]
            liczba_punktow = item[17]
            nazwa_bloku = item[19]
            zrodlo_pytania = item[20]
            komentarz = item[22]
            
            if media_file:
                
                if media_file.endswith('.wmv'):
                    # wrong file names in the XLSX for category A
                    media_file = media_file[:-4] + '.mp4'
                    
                output_media_file = os.path.join(output, "media", media_file)
                
                if not os.path.exists(output_media_file):
                    input_media_file = os.path.join(os.path.dirname(input), "media", media_file)
                    if not os.path.exists(input_media_file):
                        raise Exception("Missing media file: " + input_media_file)
                    # Windows command
                    os.system('copy "{0}" "{1}"'.format(input_media_file, output_media_file))
                    if not os.path.exists(output_media_file):
                        raise Exception("Cannot copy media file: " + input_media_file)
            question = '<div class="pytanie">' + pytanie + '</div>'
            if media_file:
                question += '<div class="media">'
                if media_file.endswith('.mp4'):
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
            outfile.write(template.format(question, answer, nazwa_bloku, liczba_punktow, komentarz, zrodlo_pytania))
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
    args = parser.parse_args()
    return args

if __name__ == '__main__':
    args = parse_arguments()
    try:
       sys.exit(main(input = args.input, media = args.media, output = args.output, category = args.category))
    except Exception as exc:
        print(exc, file=sys.stderr, end="\n")
        sys.exit(1)
    