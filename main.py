from pptx import Presentation
from pptx.util import Cm
from pptx.chart.data import ChartData, XyChartData, CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_TICK_MARK
from pptx.util import Pt
import json
from PIL import Image

pres_name = "default_name.pptx"
pres = Presentation()


def config_read():
    with open("..\\config\\config.json") as settings_json:
        ppt_configs = json.load(settings_json)
    layout = "text"
    return ppt_configs["presentation"]



def create_title_slide():
    pass


# title text list picture plot
def slide_creator(slide_config):
    left = Cm(3)
    top = Cm(3.5)

    title_type= slide_config["type"]
    title_text = slide_config["title"]
    content = slide_config["content"]
    if title_type == "title":
        #to create a title slide
        slide_type = 0
    elif title_type == "list":
        #to create a slide with list
        slide_type = 1
    else:
        #to create a slide with only a title and empty space
        slide_type = 5

    new_slide = pres.slide_layouts[slide_type]
    slide = pres.slides.add_slide(new_slide)
    slide.shapes.title.text = title_text

    if title_type == "title":
        slide.placeholders[1].text = content
    elif title_type == "list":
        pass
    elif title_type == "text":
        pass
    elif title_type == "picture":
        pass
    elif title_type == "plot":
        pass
    #image_path = "..\\images\\" + slide_config["content"]

    #list
    """level = 1-1
    list_lvl1 = slide.placeholders[1]
    list_lvl1.level = level
    list_lvl1.text = "lvl1"

    list_lvl2 = list_lvl1.text_frame.add_paragraph()
    list_lvl2.level = 1
    list_lvl2.text = "lvl2" """


    #image
    #slide.shapes.add_picture(image_path,image_left,image_top)

    #chart
    chart_data = XyChartData()
    series = chart_data.add_series("default")
    series.add_data_point(1, 2)
    series.add_data_point(3, 5)
    series.add_data_point(5, 3.1)
    chart_width, chart_length = Cm(16.25), Cm(12.5)
    chart =slide.shapes.add_chart(
        XL_CHART_TYPE.XY_SCATTER_LINES, left, top, chart_width, chart_length, chart_data
    ).chart
    x_axis = chart.category_axis
    x_axis.axis_title.text_frame.text = "xtext"
    y_axis = chart.value_axis
    y_axis.axis_title.text_frame.text = "ytext"


if __name__ == "__main__":
    slides_conf = config_read()

    for slide_conf in slides_conf:
        slide_creator(slide_conf)

    #pres_name = input("Presentation name: ") + ".pptx"
    pres.save(pres_name)
