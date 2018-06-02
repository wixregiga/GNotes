from pptx import Presentation


class Analyzer:
    """
        I want to pass in a path to the powerpoint
        Analyzer analyzes the powerpoint and spits out
        an Analysis.
    """

    def __init__(self, ppt_path):
        self.ppt_path = ppt_path
        self.presentation = Presentation(ppt_path)
        self.analysis = None
        self.title_slides = {}
        self.list_titles = []
        self.list_slides = []

        # Call the title slides function
        self.get_title_slides()

    def get_titles(self):
        for k, v in self.title_slides.items():
            self.list_titles.append(k)
        return self.list_titles

    def get_slides(self):
        for idx, slide in enumerate(self.presentation.slides):
            self.list_slides.append(slide)
        return self.list_slides

    def get_slide(self, id):
        return self.list_slides[id]

    def get_title_slides(self):
        for idx, slide in enumerate(self.presentation.slides):
            title_placeholder = slide.shapes.title
            # self.list_slides.append(slide.slide_id)  # adds all slides to list_slides
            if title_placeholder is not None:
                title = title_placeholder.text
                if title not in self.title_slides.keys():
                    self.title_slides[title] = [slide.slide_id]
                    self.list_titles.append(title)  # adds the title to the list_titles attribute
                elif slide.slide_id not in self.title_slides.get(title):
                    self.title_slides[title].append(slide.slide_id)
        return self.title_slides

    def get_slide_ids(self):
        ids = []
        for slide in self.presentation.slides:
            ids.append(slide.slide_id)
        return ids

    def get_slides_by_title(self, title):
        slides_by_title = []
        val = self.title_slides[title]
        if val is None:
            return
        else:
            for sID in val:
                slides_by_title.append(sID)
            return slides_by_title


if __name__ == '__main__':
    a = Analyzer('C:/Users/FlamingSword/Desktop/Python.pptx')
    a.get_title_slides()
    list_title_string = ''
    for title in a.list_titles:
        print(title)
        # list_title_string += '{title}\n'.format(title)
    # print(list_title_string)
