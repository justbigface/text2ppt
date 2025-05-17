def create_pptx(slides_data, output_path="output.pptx"):
    from pptx import Presentation
    prs = Presentation()
    prs.save(output_path)
    return output_path
