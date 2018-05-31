from pptx import Presentation

class Analyzer:
    '''
        I want to pass in a path to the powerpoint
        Analyzer analyzes the powerpoint and spits out
        an Analysis.
        '''

    def __init__(self, ppt_path):
        self.ppt_path = ppt_path
        self.presentation = Presentation(ppt_path)
        self.analysis = None

    def get_titles(self):
        titles = {}

        for idx, slide in enumerate(self.presentation.slides):
            title_placeholder = slide.shapes.title
            if (title_placeholder is not None):
                title = title_placeholder.text
                if title not in titles.keys():
                    #titles[title].append(slide.slide_id)
                    titles[title] = [slide.slide_id]
                elif (slide.slide_id not in titles.get(title)):
                    titles[title].append(slide.slide_id)

        return titles

class Analysis:
    '''
        ANALYSIS

        - power_point_path:
            The path to the powerpoint
        - slides:
            number of slides in the powerpoint mapped to id
        - titles: (Dictionary of title keys & slide_id numbers
            has every title in the powerpoint
        '''

    def __init__(self):
        self.power_point_path = ''
