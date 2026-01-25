import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import io
from datetime import datetime
from uuid import uuid4

# Set page to wide mode
st.set_page_config(page_title="Report Generator", layout="wide")

st.title("Field Inspection Report Generator")
st.markdown("---")

# --- 1. SESSION STATE SETUP ---
if "report_items" not in st.session_state:
    st.session_state.report_items = []
if "generated_ppt_binary" not in st.session_state:
    st.session_state.generated_ppt_binary = None
if "generated_filename" not in st.session_state:
    st.session_state.generated_filename = ""
if "uploader_id" not in st.session_state:
    st.session_state.uploader_id = 0

# Ensure each item has a stable id (important for inline widget keys + reorder)
for item in st.session_state.report_items:
    if "id" not in item:
        item["id"] = uuid4().hex

# --- 2. CALLBACK FUNCTIONS ---
def add_entry_callback():
    uploader_key = f"uploader_{st.session_state.uploader_id}"
    uploaded_file = st.session_state.get(uploader_key)
    description = st.session_state.get("entry_desc", "")
    category = st.session_state.get("cat_selector", "Exterior")
    custom_cat = (st.session_state.get("custom_cat_input") or "").strip()

    final_cat = category
    if category == "Other..." and custom_cat:
        final_cat = custom_cat
    elif category == "Other..." and not custom_cat:
        final_cat = "Other"

    if uploaded_file and description.strip():
        st.session_state.report_items.append({
            "id": uuid4().hex,
            "category": final_cat,
            "text": description.strip(),
            "image": uploaded_file
        })

        # Reset Inputs
        st.session_state["entry_desc"] = ""
        st.session_state.uploader_id += 1
        st.session_state.generated_ppt_binary = None
    else:
        st.error("Please provide both an image and a description.")

def delete_item_callback(index):
    st.session_state.report_items.pop(index)
    st.session_state.generated_ppt_binary = None

# ---- Inline edit helpers ----
def update_item_text(item_id):
    key = f"desc_{item_id}"
    for it in st.session_state.report_items:
        if it["id"] == item_id:
            it["text"] = (st.session_state.get(key) or "").strip()
            break
    st.session_state.generated_ppt_binary = None

def update_item_category(item_id):
    sel_key = f"cat_sel_{item_id}"
    other_key = f"cat_other_{item_id}"

    selected = st.session_state.get(sel_key, "Exterior")
    custom = (st.session_state.get(other_key) or "").strip()

    final_cat = selected
    if selected == "Other..." and custom:
        final_cat = custom
    elif selected == "Other..." and not custom:
        final_cat = "Other"

    for it in st.session_state.report_items:
        if it["id"] == item_id:
            it["category"] = final_cat
            break
    st.session_state.generated_ppt_binary = None

def update_item_image(item_id):
    img_key = f"img_{item_id}"
    uploaded = st.session_state.get(img_key)
    if uploaded:
        for it in st.session_state.report_items:
            if it["id"] == item_id:
                it["image"] = uploaded
                break
        st.session_state.generated_ppt_binary = None

# ---- Reorder helpers ----
def move_item(from_index, to_index):
    items = st.session_state.report_items
    if from_index < 0 or from_index >= len(items):
        return
    if to_index < 0 or to_index >= len(items):
        return
    item = items.pop(from_index)
    items.insert(to_index, item)
    st.session_state.generated_ppt_binary = None

def move_up(index):
    if index > 0:
        move_item(index, index - 1)

def move_down(index):
    if index < len(st.session_state.report_items) - 1:
        move_item(index, index + 1)

def move_top(index):
    if index > 0:
        move_item(index, 0)

def move_bottom(index):
    last = len(st.session_state.report_items) - 1
    if index < last:
        move_item(index, last)

# --- 3. SETTINGS SIDEBAR ---
with st.sidebar:
    st.header("Report Settings")
    report_title = st.text_input("Report Title", "Field Inspection Report")
    date_option = st.selectbox(
        "Date Format",
        ["Month & Year", "Date Only (MM-DD-YYYY)", "Date & Time", "Custom Text"]
    )

    report_subtitle = ""
    filename_suffix = ""

    if date_option == "Custom Text":
        report_subtitle = st.text_input("Subtitle Text", "January 2026")
        filename_suffix = report_subtitle.replace(" ", "_").replace("/", "-")
    else:
        selected_date = st.date_input("Select Date", datetime.now())
        if date_option == "Month & Year":
            report_subtitle = selected_date.strftime("%B %Y")
            filename_suffix = selected_date.strftime("%b_%Y")
        elif date_option == "Date Only (MM-DD-YYYY)":
            report_subtitle = selected_date.strftime("%m-%d-%Y")
            filename_suffix = selected_date.strftime("%m-%d-%Y")
        elif date_option == "Date & Time":
            selected_time = st.time_input("Select Time", datetime.now().time())
            final_dt = datetime.combine(selected_date, selected_time)
            report_subtitle = final_dt.strftime("%m-%d-%Y %H:%M")
            filename_suffix = final_dt.strftime("%m-%d-%Y_%H%M")

    st.divider()
    st.caption("**Preview:**")
    st.info(f"{report_subtitle}")
    clean_title = report_title.replace(" ", "_")
    final_filename = f"{clean_title}_{filename_suffix}.pptx"
    st.caption(f"**Filename:** {final_filename}")

# --- 4. BATCH UPLOAD ---
with st.expander("Batch Upload (Add Multiple Images)", expanded=False):
    st.write("Select all images in your folder and drag them here.")
    batch_files = st.file_uploader(
        "Select Multiple Images",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True
    )

    if st.button("Add All Batch Images", type="primary"):
        if batch_files:
            for file in batch_files:
                st.session_state.report_items.append({
                    "id": uuid4().hex,
                    "category": "Exterior",
                    "text": "",
                    "image": file
                })
            st.session_state.generated_ppt_binary = None
            st.success(f"Added {len(batch_files)} images! Scroll down to edit.")
        else:
            st.warning("No files selected.")

# --- 5. SINGLE ENTRY (QUICK ADD) ---
with st.container():
    st.markdown("### Add Single Entry")
    c1, c2 = st.columns([1, 1])

    with c1:
        standard_options = ["Exterior", "Interior"]
        st.selectbox("Category", standard_options + ["Other..."], key="cat_selector")

        if st.session_state.get("cat_selector") == "Other...":
            st.text_input("Enter Custom Category", key="custom_cat_input")

        st.text_area(
            "Description",
            height=150,
            placeholder="Enter observation here...",
            key="entry_desc"
        )

    with c2:
        dynamic_key = f"uploader_{st.session_state.uploader_id}"
        st.file_uploader("Upload Image (Single)", type=["png", "jpg", "jpeg"], key=dynamic_key)

    st.button("Add Entry", type="primary", on_click=add_entry_callback)

# --- 6. CURRENT ENTRIES (CLEAN CARDS + INLINE EDIT + REORDER) ---
if st.session_state.report_items:
    st.markdown("---")
    st.subheader(f"Current Entries ({len(st.session_state.report_items)})")
    st.caption("Shown in slide order (top = Slide 1). Use arrows to reorder. Edit everything inline.")

    for i, item in enumerate(st.session_state.report_items):
        item_id = item["id"]

        # Header "card"
        st.markdown(
            f"""
            <div style="
                padding: 14px 16px;
                border-radius: 12px;
                border: 1px solid #e6e6e6;
                background: #f7f8fa;
                margin-bottom: 10px;
            ">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:800; font-size:16px; color:#111;">
                        Slide {i+1}
                    </div>
                    <div style="
                        background:#111827;
                        color:#fff;
                        padding:4px 10px;
                        border-radius:999px;
                        font-size:12px;
                        font-weight:800;
                    ">
                        {item["category"]}
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        col_img, col_fields, col_actions = st.columns([2, 6, 2])

        with col_img:
            st.image(item["image"], use_container_width=True)
            st.file_uploader(
                "Replace image",
                type=["png", "jpg", "jpeg"],
                key=f"img_{item_id}",
                on_change=update_item_image,
                args=(item_id,),
                label_visibility="collapsed",
            )
            st.caption("Drop a new image to replace")

        with col_fields:
            # Category select
            standard = ["Exterior", "Interior"]
            current_cat = item["category"]
            is_standard = current_cat in standard
            default_ix = standard.index(current_cat) if is_standard else len(standard)

            st.selectbox(
                "Category",
                standard + ["Other..."],
                index=default_ix,
                key=f"cat_sel_{item_id}",
                on_change=update_item_category,
                args=(item_id,),
            )

            if st.session_state.get(f"cat_sel_{item_id}") == "Other...":
                st.text_input(
                    "Custom category",
                    value="" if is_standard else current_cat,
                    key=f"cat_other_{item_id}",
                    on_change=update_item_category,
                    args=(item_id,),
                )

            # Description
            st.text_area(
                "Description",
                value=item.get("text", ""),
                height=140,
                key=f"desc_{item_id}",
                on_change=update_item_text,
                args=(item_id,),
                placeholder="Type the observation here...",
            )

            if (item.get("text", "") or "").strip() == "":
                st.markdown(
                    """
                    <div style="
                        display:inline-block;
                        background:#fff3cd;
                        color:#664d03;
                        padding:4px 10px;
                        border-radius:999px;
                        font-size:12px;
                        font-weight:800;
                        border:1px solid #ffe69c;
                    ">
                        Missing description
                    </div>
                    """,
                    unsafe_allow_html=True
                )

        with col_actions:
            st.write("")
            st.write("")
            a1, a2 = st.columns(2)
            with a1:
                st.button("Top", key=f"top_{item_id}", on_click=move_top, args=(i,), use_container_width=True, disabled=(i == 0))
            with a2:
                st.button("Up", key=f"up_{item_id}", on_click=move_up, args=(i,), use_container_width=True, disabled=(i == 0))

            b1, b2 = st.columns(2)
            with b1:
                st.button("Down", key=f"down_{item_id}", on_click=move_down, args=(i,), use_container_width=True, disabled=(i == len(st.session_state.report_items)-1))
            with b2:
                st.button("Bottom", key=f"bot_{item_id}", on_click=move_bottom, args=(i,), use_container_width=True, disabled=(i == len(st.session_state.report_items)-1))

            st.divider()
            st.button("Delete", key=f"del_{item_id}", on_click=delete_item_callback, args=(i,), use_container_width=True)

        st.divider()

# --- 7. PPT GENERATION LOGIC ---
if st.session_state.report_items:
    if st.session_state.generated_ppt_binary is None:
        if st.button("Generate Report", type="primary", use_container_width=True):
            prs = Presentation()

            # Title slide
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = report_title
            slide.placeholders[1].text = report_subtitle

            # Layout constants
            SLIDE_WIDTH = Inches(10)
            SLIDE_HEIGHT = Inches(7.5)
            MARGIN_X = Inches(0.5)
            TOP_Y = Inches(0.8)
            GAP = Inches(0.2)
            COL_WIDTH = Inches(4.4)
            HEADER_HEIGHT = Inches(0.8)
            BODY_HEIGHT = Inches(5.4)
            IMG_BLOCK_HEIGHT = HEADER_HEIGHT + BODY_HEIGHT

            for index, item in enumerate(st.session_state.report_items):
                slide = prs.slides.add_slide(prs.slide_layouts[6])

                # Background
                background = slide.background
                background.fill.solid()
                background.fill.fore_color.rgb = RGBColor(200, 210, 215)

                # Header (category)
                header = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, MARGIN_X, TOP_Y, COL_WIDTH, HEADER_HEIGHT
                )
                header.fill.solid()
                header.fill.fore_color.rgb = RGBColor(176, 196, 222)
                header.line.color.rgb = RGBColor(0, 0, 0)
                header.text = item["category"]

                p = header.text_frame.paragraphs[0]
                p.font.bold = True
                p.font.size = Pt(26)
                p.font.color.rgb = RGBColor(0, 0, 0)
                p.alignment = PP_ALIGN.LEFT
                header.text_frame.margin_left = Inches(0.2)
                header.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                # Description box
                desc_box = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, MARGIN_X, TOP_Y + HEADER_HEIGHT, COL_WIDTH, BODY_HEIGHT
                )
                desc_box.fill.solid()
                desc_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
                desc_box.line.color.rgb = RGBColor(0, 0, 0)

                tf = desc_box.text_frame
                tf.clear()
                tf.text = item.get("text", "")
                tf.vertical_anchor = MSO_ANCHOR.TOP
                tf.margin_top = Inches(0.2)
                tf.margin_left = Inches(0.2)
                tf.word_wrap = True

                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    p.font.bold = False
                    p.font.size = Pt(20)
                    p.font.color.rgb = RGBColor(0, 0, 0)
                    p.alignment = PP_ALIGN.LEFT

                # Image
                img_x = MARGIN_X + COL_WIDTH + GAP
                try:
                    item["image"].seek(0)
                except Exception:
                    pass

                pic = slide.shapes.add_picture(
                    item["image"], img_x, TOP_Y, width=COL_WIDTH, height=IMG_BLOCK_HEIGHT
                )
                pic.line.color.rgb = RGBColor(0, 0, 0)
                pic.line.width = Pt(1)

                # Footer
                footer_y = SLIDE_HEIGHT - Inches(0.5)
                footer_box = slide.shapes.add_textbox(MARGIN_X, footer_y, Inches(6), Inches(0.5))
                fp = footer_box.text_frame.paragraphs[0]
                fp.text = report_title
                fp.font.size = Pt(10)
                fp.font.color.rgb = RGBColor(80, 80, 80)

                page_box = slide.shapes.add_textbox(
                    SLIDE_WIDTH - MARGIN_X - Inches(2), footer_y, Inches(2), Inches(0.5)
                )
                pp = page_box.text_frame.paragraphs[0]
                pp.text = f"Page {index + 1}"
                pp.font.size = Pt(10)
                pp.font.color.rgb = RGBColor(80, 80, 80)
                pp.alignment = PP_ALIGN.RIGHT

            binary_output = io.BytesIO()
            prs.save(binary_output)
            binary_output.seek(0)

            st.session_state.generated_ppt_binary = binary_output
            st.session_state.generated_filename = final_filename
            st.rerun()
    else:
        st.download_button(
            label=f"Download {st.session_state.generated_filename}",
            data=st.session_state.generated_ppt_binary,
            file_name=st.session_state.generated_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
            use_container_width=True,
        )

        if st.button("Reset / Start New Report", use_container_width=True):
            st.session_state.report_items = []
            st.session_state.generated_ppt_binary = None
            st.session_state.uploader_id += 1
            st.rerun()
