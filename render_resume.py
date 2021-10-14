from docxtpl import DocxTemplate, RichText
import pprint

# Import the template
template = DocxTemplate('resume_template.docx')


# define the context dictionary to feed into the template
context = {'personName' : 'John Doe',
'location' : 'New York City, NY 11200',
'phone' : '123456789',
'email' : 'john@doe.com',
'linkedin' : 'linkedin/johndoe',
'website' : 'johndoe.com',
'citizenship' : 'USA',
'introduction' : 'I\'m a really big deal.'}

jobs = [
{'title' : 'Software Engineer',
'company' : 'Google',
'location' : 'Mountain View, California',
'startDate' : 'March, 2020',
'endDate' : 'March, 2021',
'responsibilities' : ['cleaning toilets','scrubbing the deck','professional networking']},
{'title' : 'Software Developer',
'company' : 'Facebook',
'location' : 'Menlo Park, California',
'startDate' : 'March, 2018',
'endDate' : 'March, 2019',
'responsibilities' : ['cleaning toilets','scrubbing the deck','professional networking']}
]
context['jobs'] = jobs

degrees = [
{'degreeName': 'MSc Computer Science',
'college' : 'Hardvard',
'location' : 'USA',
'startYear' : '2020',
'endYear' : '2021',
'GPA' : '4.0'},
{'degreeName': 'BSc Computer Science',
'college' : 'Hardvard',
'location' : 'USA',
'startYear' : '2016',
'endYear' : '2020',
'GPA' : '4.0'}
]
context['degrees'] = degrees

# Previous projects, including links and code
projectList = [
{'description': 'Big open source python project',
'link': 'gmail.com',
'code': 'github.com/johndoe/gmail-codebase',},
{'description': 'Database engine that is very fast',
'link': 'postgres.com',},
{'description': 'other cool stuff that gets you jobs',
'code': 'github.com/johndoe/cool-stuff'},
{'description': 'Stuff without links',}
]
context['projectList'] = projectList

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
context['projects'] = renderProjectLinks(projectList, template)

# Dictionary of skills keywords
skills = {'Programming' : ['C++', 'Python', 'Rust'],
'Data' : ['MySQL', 'Excel']}
context['skillDict'] = skills

# For testing, pretty print the dictionary
pp = pprint.PrettyPrinter(indent=4)
pp.pprint(context)

#Render automated report
template.render(context)
template.save('generated_resume.docx')
