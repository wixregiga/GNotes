from pptx import Presentation
from ganalyze.analyzer import Analyzer
# dir_path = os.path.dirname(os.path.realpath(__file__))
# print(dir_path)

# presentation_path = 'C‪:\\Users\\FlamingSword\\Downloads\\Schoolhouse\\Block 03 - Scripting\\Slides\\Python.pptx'
# prs = Presentation(presentation_path)


presentation_path = 'C:/Users/FlamingSword/Desktop/Python.pptx'
my_analyzer = Analyzer(presentation_path)

titles = my_analyzer.get_titles()

for key in titles:
   print('TITLE: {},   SLIDES:   {}'.format(key, str(titles[key])))

# for k, v in enumerate(titles.items()):
#     print('TITLE: {},   SLIDES:   {}'.format(k, str(v)))



# prs = Presentation(presentation_path)
'''
    X   1. Print numbaer of slides
    x   2. Print List of all titles along w/ slide number
    -   3. 
'''

# 1. Print number of slides
# num_slides = len(prs.slides)
# print(num_slides)

# 2. Print list of all titles along w/ slide number
# for idx, slide in enumerate(prs.slides):
#     slide.
#     title = slide.shapes.title
#     if title is not None:
#         print('%d %s' % (idx, title.text))

# slf = []
# for i in range(2):
#     slf.append(prs.slides[i])
# for slide in enumerate(prs.slides)
# 2. Print List of all (first 5) titles along with their corresponding slide number

# for slide in slf:
#     for shape in slide.placeholders:
#         print(len(shape))
#         print(shape.name)
#         print('%d %s' % (shape.placeholder_format.idx, shape.name))

# for shape in prs.slides[2].placeholders:
#     print('%d %s' % (shape.placeholder_format.idx, shape.name))

#
# for shape in slide.placeholders:
# ...     print('%d %s' % (shape.placeholder_format.idx, shape.name))


# for shape in sl.shapes:
#     if not shape.has_text_frame:
#         continue
#     for paragraph in shape.text_frame.paragraphs:
#         for run in paragraph.runs:
#             text_runs.append(run.text)
#
# for text in text_runs:
#     print(text)