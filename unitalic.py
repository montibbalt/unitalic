import argparse
import docx

def unitalic(doc):
    for pp in doc.paragraphs:
        in_quote = False
        for run in pp.runs:
            if run.text.startswith(('"', '“')):
                in_quote = True

            if in_quote:
                run.italic = False

            if in_quote and run.text.endswith(('"', '”')):
                in_quote = False

def main(args):
    doc = docx.Document(args.input)
    unitalic(doc)
    doc.save(args.output)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(prog="UnItalic", description="removes italics from quotations in Word documents")
    parser.add_argument('-i', '--input', required=True)
    parser.add_argument('-o', '--output', required=True)
    args = parser.parse_args()
    main(args)