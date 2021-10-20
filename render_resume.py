from docxtpl import DocxTemplate, RichText
import pprint
import json
import re
import argparse

# render the projectList list into RichText with embedded links
def renderProjectLinks(projectList, template):
    projectListCombined = []
    for project in projectList:
        richTextString = RichText(project['description'])
        if 'link' in project and 'code' in project:
            richTextString.add(' - ')
            richTextString.add('link',url_id=template.build_url_id(project['link']) )
            richTextString.add(', ')
            richTextString.add('code',url_id=template.build_url_id(project['code']) )
        elif 'link' in project:
            richTextString.add(' - ')
            richTextString.add('link',url_id=template.build_url_id(project['link']) )
            richTextString.add('.')
        elif 'code' in project:
            richTextString.add(' - ')
            richTextString.add('code',url_id=template.build_url_id(project['code']) )
            richTextString.add('.')
        else:
            richTextString.add('.')
        projectListCombined.append(richTextString)
    return projectListCombined

def dictToDocx(context):
    # Import the template
    template = DocxTemplate('resume_template.docx')

    # For testing, pretty print the dictionary
    # pp = pprint.PrettyPrinter(indent=4)
    # pp.pprint(context)

    # generate filename
    genFilename = lambda theString, extension : re.sub( '[^0-9a-zA-Z ]+','',theString.lower()).replace(' ', '_')+'.'+extension

    # Export to json file
    # JSONFile=genFilename(context['intro']['personName'],'json')
    # with open(JSONFile,'w') as resumeJSONFile:
    #     json.dump(context, resumeJSONFile, indent=4)

    # render links as click-able hyperlinks using the richTextString class
    context['projects'] = renderProjectLinks(context['projectList'], template)

    #Render automated report
    docxFile=genFilename(context['intro']['personName'],'docx')
    print("Generating resume document: {}".format(docxFile))
    template.render(context)
    template.save(docxFile)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Produce a ATS-friendly docx resume.')
    parser.add_argument('-j','--json', type=open, nargs=1, help='path to .json file of data')

    args = parser.parse_args()
    if args.json is not None:
        with open(args.json) as f:
            data = json.load(f)
        dictToDocx(data)
    else:
        print('Generating sample docx document from sample_data.json')
        with open('sample_data.json') as f:
            data = json.load(f)
        dictToDocx(data)
