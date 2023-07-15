from pptx import Presentation
import json
import slide_creator


def config_read():
    with open("..\\config\\config.json") as settings_json:
        ppt_configs = json.load(settings_json)
    return ppt_configs["presentation"]


def slide_maker(slide_config):
    title_type = slide_config["type"]
    title_text = slide_config["title"]
    content = slide_config["content"]
    if title_type == "title":
        slide_creator.create_title_slide(pres, title_text, content)
    elif title_type == "list":
        slide_creator.create_list_slide(pres, title_text, content)
    elif title_type == "text":
        slide_creator.create_text_slide(pres, title_text, content)
    elif title_type == "picture":
        slide_creator.create_picture_slide(pres, title_text, content)
    elif title_type == "plot":
        slide_creator.create_plot_slide(pres, title_text, content, slide_config["configuration"])


if __name__ == "__main__":
    pres_name = "..\\presentations\\default_name.pptx"
    # pres_name = "..\\presentations\\" + input("Presentation name: ") + ".pptx"
    pres = Presentation()
    slides_conf = config_read()
    for slide_conf in slides_conf:
        slide_maker(slide_conf)
    pres.save(pres_name)
