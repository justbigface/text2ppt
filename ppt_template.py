def create_pptx(slides, output_path="output.pptx"):
    from pptx import Presentation
    pres = Presentation()
    pres.save(output_path)
    return output_path
