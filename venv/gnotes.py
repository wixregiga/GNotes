from ganalyze.analyzer import Analyzer

presentation_path = 'C:/Users/FlamingSword/Desktop/Python.pptx'
gnote_output_path = 'C:/Users/FlamingSword/Desktop/gnote_output.txt'

my_analyzer = Analyzer(presentation_path)


def test_title_slides(analyzer):
    titles = analyzer.get_title_slides()
    sholder = ''
    for key in titles:
        sholder += 'TITLE: {},   SLIDES:   {}\n'.format(key, str(titles[key]))
    return sholder


def test_get_all_slide_ids(analyzer):
    return str(analyzer.get_slide_ids())


def test_get_all_slides(analyzer):
    return str(analyzer.get_slides())


def test_get_slides_by_title(analyzer):
    slides_by_title = analyzer.get_slides_by_title('Mathematical Operations')
    sholder = ''
    for slide in slides_by_title:
        sholder += '{} \n'.format(str(slide))
    return sholder


def test_get_slide(analyzer, id):
    return str(analyzer.slide(id))


def iprint(mfunc, analyzer):
    name = mfunc.__name__.upper()

    print('\n++++++++++++++++++++++ NOW RUNNING TEST: {}      \n\n'.format(name))

    print(mfunc(analyzer))

    print('\n----------------------- FINISHED RUNNING TEST: {} \n\n\n'.format(name))

def write_ppt_to_file():
    text_runs = []
    gnote_output = open(gnote_output_path, 'w')

    for slide in my_analyzer.presentation.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text_runs.append(run.text)
    for text in text_runs:
        # print('{}\n'.format(text))
        gnote_output.write('{}\n'.format(text))

    gnote_output.close()


if __name__ == '__main__':
    # iprint(test_title_slides, my_analyzer)
    # iprint(test_get_all_slides, my_analyzer)
    # iprint(test_get_all_slide_ids, my_analyzer)
    # iprint(test_get_slides_by_title, my_analyzer)
    write_ppt_to_file()





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
