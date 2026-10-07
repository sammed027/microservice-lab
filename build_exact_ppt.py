import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Color Palette matching friend's PPT exact theme
    NAVY_DARK = RGBColor(23, 35, 61)       # #17233d
    NAVY_HEADER = RGBColor(23, 35, 61)     # Header banner
    BG_LIGHT = RGBColor(244, 246, 249)     # Slide body bg #f4f6f9
    TEXT_WHITE = RGBColor(255, 255, 255)
    TEXT_DARK = RGBColor(31, 41, 55)       # Body text #1f2937
    TEXT_MUTED = RGBColor(100, 116, 139)   # Subtitle text #64748b
    TEAL_ACCENT = RGBColor(0, 150, 136)    # Teal border #009688
    ORANGE_ACCENT = RGBColor(230, 126, 34) # Orange box/bar #e67e22
    CARD_BG = RGBColor(255, 255, 255)      # White card bg

    def set_slide_background(slide, color):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()

    def add_top_banner(slide, title_text):
        # Header banner background
        banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.1))
        banner.fill.solid()
        banner.fill.fore_color.rgb = NAVY_HEADER
        banner.line.fill.background()

        # Teal Rule Line under banner
        rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.1), Inches(13.333), Inches(0.06))
        rule.fill.solid()
        rule.fill.fore_color.rgb = TEAL_ACCENT
        rule.line.fill.background()

        # Title Text
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.25), Inches(11.7), Inches(0.7))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = TEXT_WHITE

    def add_footer(slide):
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(7.0), Inches(11.7), Inches(0.4))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = "Gym Management Microservices | Cloud Computing Lab Evaluation"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_MUTED

    # -------------------------------------------------------------
    # SLIDE 1: TITLE SLIDE (Matching Screenshot 1)
    # -------------------------------------------------------------
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1, NAVY_DARK)

    # Orange Accent Vertical Bar
    orange_bar = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.1), Inches(0.15), Inches(2.6))
    orange_bar.fill.solid()
    orange_bar.fill.fore_color.rgb = ORANGE_ACCENT
    orange_bar.line.fill.background()

    # Title Text Box
    tb1 = slide1.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(11.3), Inches(3.2))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p_main = tf1.paragraphs[0]
    p_main.text = "Build, Deploy and Analyze a Containerized\nMicroservice Application Under Varying Workloads"
    p_main.font.size = Pt(32)
    p_main.font.bold = True
    p_main.font.color.rgb = TEXT_WHITE

    p_sub = tf1.add_paragraph()
    p_sub.text = "Gym Management Microservices"
    p_sub.font.size = Pt(22)
    p_sub.font.color.rgb = RGBColor(147, 197, 253) # Light cyan blue
    p_sub.space_before = Pt(20)

    p_eval = tf1.add_paragraph()
    p_eval.text = "Cloud Computing Lab Evaluation"
    p_eval.font.size = Pt(18)
    p_eval.font.color.rgb = RGBColor(203, 213, 225)
    p_eval.space_before = Pt(8)

    # Presentation Footer Info
    tb_present = slide1.shapes.add_textbox(Inches(1.2), Inches(6.2), Inches(11.0), Inches(0.6))
    tf_present = tb_present.text_frame
    p_pres = tf_present.paragraphs[0]
    p_pres.text = "Presented by: Student Team   |   KLE Technological University, BVB Campus, Hubli"
    p_pres.font.size = Pt(12)
    p_pres.font.color.rgb = RGBColor(203, 213, 225)

    # -------------------------------------------------------------
    # SLIDE 2: AIM AND OBJECTIVES (Matching Screenshot 2)
    # -------------------------------------------------------------
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2, BG_LIGHT)
    add_top_banner(slide2, "Aim and Objectives")

    # Main Aim Summary Paragraph
    tb_aim = slide2.shapes.add_textbox(Inches(0.8), Inches(1.35), Inches(11.7), Inches(0.8))
    tf_aim = tb_aim.text_frame
    tf_aim.word_wrap = True
    p_aim = tf_aim.paragraphs[0]
    p_aim.text = "To build a microservice application with three independent services, containerize and deploy it with Docker and Docker Compose, generate varying workloads, monitor resource use, and analyze performance."
    p_aim.font.size = Pt(15)
    p_aim.font.color.rgb = TEXT_DARK

    # 4 Cards Grid
    card_data = [
        ("Develop", "Three independent FastAPI microservices with REST APIs"),
        ("Containerize", "One Dockerfile and one Docker image per service"),
        ("Deploy", "Docker Compose runs all three containers on one network"),
        ("Analyze", "Load test at 1, 2, 4, 8, 16 concurrent requests and monitor CPU and memory")
    ]

    card_width = Inches(2.7)
    card_height = Inches(3.2)
    start_left = Inches(0.8)
    spacing = Inches(0.3)

    for i, (title, desc) in enumerate(card_data):
        left_pos = start_left + i * (card_width + spacing)
        top_pos = Inches(2.4)

        # White Card Shape
        card = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_pos, top_pos, card_width, card_height)
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = RGBColor(226, 232, 240)

        # Teal Vertical Accent Bar on Left
        t_bar = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_pos, top_pos, Inches(0.12), card_height)
        t_bar.fill.solid()
        t_bar.fill.fore_color.rgb = TEAL_ACCENT
        t_bar.line.fill.background()

        # Text inside Card
        tb_card = slide2.shapes.add_textbox(left_pos + Inches(0.25), top_pos + Inches(0.2), card_width - Inches(0.3), card_height - Inches(0.4))
        tf_card = tb_card.text_frame
        tf_card.word_wrap = True

        p_t = tf_card.paragraphs[0]
        p_t.text = title
        p_t.font.size = Pt(18)
        p_t.font.bold = True
        p_t.font.color.rgb = NAVY_DARK

        p_d = tf_card.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(13)
        p_d.font.color.rgb = TEXT_DARK
        p_d.space_before = Pt(12)

    # Footer note (italicized)
    tb_note = slide2.shapes.add_textbox(Inches(0.8), Inches(6.1), Inches(11.7), Inches(0.5))
    tf_note = tb_note.text_frame
    p_n = tf_note.paragraphs[0]
    p_n.text = "Core challenge: moving from one monolith to independent, networked containers while finding the performance bottleneck."
    p_n.font.size = Pt(13)
    p_n.font.italic = True
    p_n.font.color.rgb = TEAL_ACCENT

    add_footer(slide2)

    # -------------------------------------------------------------
    # SLIDE 3: TECHNOLOGY STACK AND ENVIRONMENT (Matching Screenshot 3)
    # -------------------------------------------------------------
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3, BG_LIGHT)
    add_top_banner(slide3, "Technology Stack and Environment")

    tech_grid = [
        ("Backend", "Python 3.11 + FastAPI"),
        ("Web server", "Uvicorn (ASGI worker)"),
        ("Containerization", "Docker 29.8.0"),
        ("Orchestration", "Docker Compose v5.5.1"),
        ("Load testing", "Python thread pool / httpx benchmark"),
        ("Test machine", "macOS, 6 CPUs, 3.8 GB for Docker")
    ]

    box_w = Inches(3.7)
    box_h = Inches(2.2)
    col_spacing = Inches(0.3)
    row_spacing = Inches(0.3)

    for i, (title, desc) in enumerate(tech_grid):
        col = i % 3
        row = i // 3

        left_pos = Inches(0.8) + col * (box_w + col_spacing)
        top_pos = Inches(1.6) + row * (box_h + row_spacing)

        # White Box
        box = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_pos, top_pos, box_w, box_h)
        box.fill.solid()
        box.fill.fore_color.rgb = CARD_BG
        box.line.color.rgb = RGBColor(226, 232, 240)

        # Teal Vertical Accent Bar
        t_bar = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_pos, top_pos, Inches(0.12), box_h)
        t_bar.fill.solid()
        t_bar.fill.fore_color.rgb = TEAL_ACCENT
        t_bar.line.fill.background()

        # Text Frame
        tb_box = slide3.shapes.add_textbox(left_pos + Inches(0.25), top_pos + Inches(0.2), box_w - Inches(0.3), box_h - Inches(0.4))
        tf_box = tb_box.text_frame
        tf_box.word_wrap = True

        p_t = tf_box.paragraphs[0]
        p_t.text = title
        p_t.font.size = Pt(16)
        p_t.font.bold = True
        p_t.font.color.rgb = NAVY_DARK

        p_d = tf_box.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(18)
        p_d.font.bold = True
        p_d.font.color.rgb = TEXT_DARK
        p_d.space_before = Pt(14)

    add_footer(slide3)

    # -------------------------------------------------------------
    # SLIDE 4: SYSTEM ARCHITECTURE (Matching Screenshot 4)
    # -------------------------------------------------------------
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4, BG_LIGHT)
    add_top_banner(slide4, "System Architecture")

    # Top Client Box
    client_box = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.2), Inches(1.5), Inches(2.9), Inches(0.8))
    client_box.fill.solid()
    client_box.fill.fore_color.rgb = RGBColor(100, 116, 139) # Slate grey
    client_box.line.fill.background()

    tf_c = client_box.text_frame
    p_c1 = tf_c.paragraphs[0]
    p_c1.text = "Client"
    p_c1.font.size = Pt(16)
    p_c1.font.bold = True
    p_c1.font.color.rgb = TEXT_WHITE
    p_c1.alignment = PP_ALIGN.CENTER

    p_c2 = tf_c.add_paragraph()
    p_c2.text = "curl / load test / web app"
    p_c2.font.size = Pt(11)
    p_c2.font.color.rgb = RGBColor(226, 232, 240)
    p_c2.alignment = PP_ALIGN.CENTER

    # Outer Dashed Container Network Box
    net_box = slide4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.5), Inches(2.7), Inches(8.3), Inches(4.0))
    net_box.fill.solid()
    net_box.fill.fore_color.rgb = RGBColor(240, 253, 250) # Light teal bg
    net_box.line.color.rgb = TEAL_ACCENT
    net_box.line.width = Pt(2)

    # Network Label
    tb_net = slide4.shapes.add_textbox(Inches(2.7), Inches(2.8), Inches(5.0), Inches(0.4))
    p_n = tb_net.text_frame.paragraphs[0]
    p_n.text = "Docker network: gym-network"
    p_n.font.size = Pt(14)
    p_n.font.bold = True
    p_n.font.color.rgb = TEAL_ACCENT

    # Central Entry Point Box: attendance-service
    att_box = slide4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.5), Inches(3.5), Inches(4.3), Inches(1.1))
    att_box.fill.solid()
    att_box.fill.fore_color.rgb = NAVY_DARK
    att_box.line.fill.background()

    tf_att = att_box.text_frame
    p_a1 = tf_att.paragraphs[0]
    p_a1.text = "attendance-service"
    p_a1.font.size = Pt(18)
    p_a1.font.bold = True
    p_a1.font.color.rgb = TEXT_WHITE
    p_a1.alignment = PP_ALIGN.CENTER

    p_a2 = tf_att.add_paragraph()
    p_a2.text = "Port 8003 (entry point)"
    p_a2.font.size = Pt(12)
    p_a2.font.color.rgb = RGBColor(203, 213, 225)
    p_a2.alignment = PP_ALIGN.CENTER

    # Bottom Left Box: member-service
    mem_box = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.0), Inches(5.3), Inches(3.4), Inches(1.1))
    mem_box.fill.solid()
    mem_box.fill.fore_color.rgb = TEAL_ACCENT
    mem_box.line.fill.background()

    tf_m = mem_box.text_frame
    p_m1 = tf_m.paragraphs[0]
    p_m1.text = "member-service"
    p_m1.font.size = Pt(17)
    p_m1.font.bold = True
    p_m1.font.color.rgb = TEXT_WHITE
    p_m1.alignment = PP_ALIGN.CENTER

    p_m2 = tf_m.add_paragraph()
    p_m2.text = "Port 8001"
    p_m2.font.size = Pt(12)
    p_m2.font.color.rgb = TEXT_WHITE
    p_m2.alignment = PP_ALIGN.CENTER

    # Bottom Right Box: membership-service
    ms_box = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(5.3), Inches(3.4), Inches(1.1))
    ms_box.fill.solid()
    ms_box.fill.fore_color.rgb = ORANGE_ACCENT
    ms_box.line.fill.background()

    tf_ms = ms_box.text_frame
    p_ms1 = tf_ms.paragraphs[0]
    p_ms1.text = "membership-service"
    p_ms1.font.size = Pt(17)
    p_ms1.font.bold = True
    p_ms1.font.color.rgb = TEXT_WHITE
    p_ms1.alignment = PP_ALIGN.CENTER

    p_ms2 = tf_ms.add_paragraph()
    p_ms2.text = "Port 8002"
    p_ms2.font.size = Pt(12)
    p_ms2.font.color.rgb = TEXT_WHITE
    p_ms2.alignment = PP_ALIGN.CENTER

    add_footer(slide4)

    # -------------------------------------------------------------
    # SLIDE 5: MICROSERVICES AND REST API ENDPOINTS (Matching Screenshot 5)
    # -------------------------------------------------------------
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5, BG_LIGHT)
    add_top_banner(slide5, "Microservices and REST API Endpoints")

    # Table
    t_shape = slide5.shapes.add_table(4, 4, Inches(0.8), Inches(1.6), Inches(11.7), Inches(4.8))
    table = t_shape.table
    table.columns[0].width = Inches(2.5)
    table.columns[1].width = Inches(1.2)
    table.columns[2].width = Inches(4.5)
    table.columns[3].width = Inches(3.5)

    headers = ["Service", "Port", "Responsibility", "Endpoints"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY_DARK
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.bold = True
        p.font.size = Pt(14)
        p.font.color.rgb = TEXT_WHITE

    rows_data = [
        ("attendance-service", "8003", "Entry point. Performs check-in by calling member and membership services.", "GET /\nPOST /attendance/checkin\nPOST /attendance/checkout/{id}\nGET /attendance/{member_id}"),
        ("member-service", "8001", "Stores and returns gym member profile records.", "GET /\nGET /members\nPOST /members\nGET /members/{id}\nPUT /members/{id}\nDELETE /members/{id}"),
        ("membership-service", "8002", "Stores membership plan tiers and validity status.", "GET /\nGET /memberships\nPOST /memberships\nGET /memberships/{member_id}\nPUT /memberships/{member_id}\nDELETE /memberships/{member_id}")
    ]

    for row_idx, rdata in enumerate(rows_data, start=1):
        for col_idx, cell_value in enumerate(rdata):
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD_BG
            p = cell.text_frame.paragraphs[0]
            p.text = cell_value
            p.font.size = Pt(12)
            p.font.color.rgb = TEXT_DARK if col_idx != 0 else NAVY_DARK
            if col_idx == 0:
                p.font.bold = True

    add_footer(slide5)

    # Save presentation
    output_filename = "Gym_Microservices_Lab_Evaluation.pptx"
    prs.save(output_filename)
    print(f"Presentation generated successfully as '{output_filename}'.")

if __name__ == "__main__":
    create_presentation()
