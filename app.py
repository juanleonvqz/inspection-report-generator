import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import io
from datetime import datetime
from uuid import uuid4
from PIL import Image

# --------------------------------------------------
# Page setup
# --------------------------------------------------
st.set_page_config(page_title="Field Inspection Report Generator", layout="wide")
st.title("Field Inspection Report Generator")
st.markdown("---")

# --------------------------------------------------
# Session state
# --------------------------------------------------
if "report_items" not in st.session_state:
    st.session_state.report_items = []
if "generated_ppt_binary" not in st.session_state:
    st.session_state.generated_ppt_binary = None
if "generated_filename" not in st.session_state:
    st.session_state.generated_filename = ""
if "uploader_id" not in st.session_state:
    st.session_state.uploader_id = 0

# Ensure each item has an id
for item in st.session_state.report_items:
    if "id" not in item:
        item["id"] = uuid4().hex

# --------------------------------------------------
# Helpers
# --------------------------------------------------
def get_image_wh(uploaded):
    try:
        uploaded.seek(0)
    except:
        pass
    img = Image.open(uploaded)
    w, h = img.size
    try:
        uploaded.seek(0)
    except:
        pass
    return w, h

# --------------------------------------------------
# Callbacks
# --------------------------------------------------
def add_entry_callback():
    uploader_key = f"uploader_{st.session_state.uploader_id}"
    uploaded_file = st.session_state.get(uploader_key)
    description = (st.session_state.get("entry_desc") or "").strip()
    category = st.session_state.get("cat_selector", "Exterior")
    custom_cat = (st.session_state.get("custom_cat_input") or "").strip()

    final_cat = category
    if category == "Other..." and custom_cat:
        final_cat = custom_cat
    elif category == "Other...":
        final_cat = "Other"

    if uploaded_file and description:
        st.session_state.report_items.append({
            "id": uuid4().hex,
            "category": final_cat,
            "text": description,
            "image": uploaded_file
        })
        st.session_state["entry_desc"] = ""
        st.session_state.uploader_id += 1
        st.session_state.generated_ppt_binary = None
    else:
        st.error("Please provide both an image and a description.")

def delete_item_callback(index):
    st.session_state.report_items.pop(index)
    st.session_state.generated_ppt_binary = None

def update_item_text(item_id):
    for it in st.session_state.report_items:
        if it["id"] == item_id:
            it["text"] = (st.session_state.get(f"desc_{item_id}") or "").strip()
            break
    st.session_state.generated_ppt_binary = None

def update_item_category(item_id):
    selected = st.session_state.get(f"cat_sel_{item_id}", "Exterior")
    custom = (st.session_state.get(f"cat_other_{item_id}") or "").strip()
    final_cat = custom if selected == "Other..." and custom else selected
    for it in st.session_state.report_items:
        if it["id"] == item_id:
            it["category"] = final_cat
            break
    st.session_state.generated_ppt_binary = None

def update_item_image(item_id):
    uploaded = st.session_state.get(f"img_{item_id}")
    if uploaded:
        for it in st.session_state.report_items:
            if it["id"] == item_id:
                it["image"] = uploaded
                break
        st.session_state.generated_ppt_binary = None

def move_item(frm, to):
    items = st.session_state.report_items
    item = items.pop(frm)
    items.insert(to, item)
    st.session_state.generated_ppt_binary = None

def move_up(i): 
    if i > 0: move_item(i, i-1)

def move_down(i): 
    if i < len(st.session_state.report_items)-1: move_item(i, i+1)

def move_top(i): 
    if i > 0: move_item(i, 0)

def move_bottom(i): 
    if i < len(st.session_state.report_items)-1: move_item(i, len(st.session_state.report_items)-1)

# --------------------------------------------------
# Sidebar
# --------------------------------------------------
with st.sidebar:
    st.header("Report Settings")
    report_title = st.text_input("Report Title", "Field Inspection Report")
    date_option = st.selectbox("Date Format", ["Month & Year", "Date Only (MM-DD-YYYY)", "Date & Time", "Custom Text"])

    if date_option == "Custom Text":
        report_subtitle = st.text_input("Subtitle Text", "January 2026")
    else:
        d = st.date_input("Select Date", datetime.now())
        if date_option == "Month & Year":
            report_subtitle = d.strftime("%B %Y")
        elif date_option == "Date Only (MM-DD-YYYY)":
            report_subtitle = d.strftime("%m-%d-%Y")
        else:
            t = st.time_input("Select Time", datetime.now().time())
            report_subtitle = datetime.combine(d, t).strftime("%m-%d-%Y %H:%M")

    final_filename = f"{report_title.replace(' ','_')}_{report_subtitle.replace(' ','_')}.pptx"
    st.info(report_subtitle)
    st.caption(final_filename)

# --------------------------------------------------
# Batch upload
# --------------------------------------------------
with st.expander("Batch Upload Images"):
    files = st.file_uploader("Select images", type=["png","jpg","jpeg"], accept_multiple_files=True)
    if st.button("Add Batch"):
        if files:
            for f in files:
                st.session_state.report_items.append({
                    "id": uuid4().hex,
                    "category": "Exterior",
                    "text": "",
                    "image": f
                })
            st.success(f"Added {len(files)} images.")
        else:
            st.warning("No images selected.")

# --------------------------------------------------
# Quick add
# --------------------------------------------------
st.subheader("Add Single Entry")
c1, c2 = st.columns(2)
with c1:
    st.selectbox("Category", ["Exterior","Interior","Other..."], key="cat_selector")
    if st.session_state.get("cat_selector") == "Other...":
        st.text_input("Custom Category", key="custom_cat_input")
    st.text_area("Description", height=120, key="entry_desc")
with c2:
    st.file_uploader("Upload image", type=["png","jpg","jpeg"], key=f"uploader_{st.session_state.uploader_id}")

st.button("Add Entry", type="primary", on_click=add_entry_callback)

# --------------------------------------------------
# Current entries
# --------------------------------------------------
if st.session_state.report_items:
    st.markdown("---")
    st.subheader("Current Entries")

    for i, item in enumerate(st.session_state.report_items):
        item_id = item["id"]

        st.markdown(f"""
        <div style="padding:12px;border-radius:12px;border:1px solid #e5e7eb;background:#f9fafb;margin-bottom:10px;">
        <b>Slide {i+1}</b> 
        <span style="float:right;background:#111827;color:white;padding:4px 10px;border-radius:999px;font-size:12px;">{item['category']}</span>
        </div>
        """, unsafe_allow_html=True)

        img, fields, actions = st.columns([2,6,2])

        with img:
            st.image(item["image"], use_container_width=True)
            st.file_uploader("Replace", type=["png","jpg","jpeg"], key=f"img_{item_id}", on_change=update_item_image, args=(item_id,), label_visibility="collapsed")

        with fields:
            base = ["Exterior","Interior"]
            cur = item["category"]
            is_std = cur in base
            idx = base.index(cur) if is_std else len(base)

            st.selectbox("Category", base+["Other..."], index=idx, key=f"cat_sel_{item_id}", on_change=update_item_category, args=(item_id,))
            if st.session_state.get(f"cat_sel_{item_id}") == "Other...":
                st.text_input("Custom", value="" if is_std else cur, key=f"cat_other_{item_id}", on_change=update_item_category, args=(item_id,))
            st.text_area("Description", value=item["text"], height=120, key=f"desc_{item_id}", on_change=update_item_text, args=(item_id,))

        with actions:
            st.button("Top", on_click=move_top, args=(i,), use_container_width=True, disabled=i==0)
            st.button("Up", on_click=move_up, args=(i,), use_container_width=True, disabled=i==0)
            st.button("Down", on_click=move_down, args=(i,), use_container_width=True, disabled=i==len(st.session_state.report_items)-1)
            st.button("Bottom", on_click=move_bottom, args=(i,), use_container_width=True, disabled=i==len(st.session_state.report_items)-1)
            st.divider()
            st.button("Delete", on_click=delete_item_callback, args=(i,), use_container_width=True)

# --------------------------------------------------
# Generate PPT
# --------------------------------------------------
if st.session_state.report_items:
    if st.session_state.generated_ppt_binary is None:
        if st.button("Generate Report", type="primary", use_container_width=True):
            prs = Presentation()

            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = report_title
            slide.placeholders[1].text = report_subtitle

            for index, item in enumerate(st.session_state.report_items):
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                bg = slide.background
                bg.fill.solid()
                bg.fill.fore_color.rgb = RGBColor(200,210,215)

                w, h = get_image_wh(item["image"])
                is_landscape = (w / h) >= 1.2 if h else False

                SLIDE_W = Inches(10)
                SLIDE_H = Inches(7.5)
                M = Inches(0.5)

                if not is_landscape:
                    COL = Inches(4.4)
                    HEAD = Inches(0.8)
                    BODY = Inches(5.4)

                    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, M, Inches(0.7), COL, HEAD)
                    header.fill.solid()
                    header.fill.fore_color.rgb = RGBColor(176,196,222)
                    header.text = item["category"]
                    header.text_frame.paragraphs[0].font.size = Pt(26)
                    header.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                    desc = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, M, Inches(0.7)+HEAD, COL, BODY)
                    desc.fill.solid()
                    desc.fill.fore_color.rgb = RGBColor(255,255,255)
                    tf = desc.text_frame
                    tf.text = item["text"]
                    tf.paragraphs[0].font.size = Pt(20)

                    item["image"].seek(0)
                    slide.shapes.add_picture(item["image"], M+COL+Inches(0.2), Inches(0.7), width=COL, height=HEAD+BODY)

                else:
                    FULL = SLIDE_W - M*2
                    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, M, Inches(0.6), FULL, Inches(0.8))
                    header.fill.solid()
                    header.fill.fore_color.rgb = RGBColor(176,196,222)
                    header.text = item["category"]
                    header.text_frame.paragraphs[0].font.size = Pt(26)

                    item["image"].seek(0)
                    slide.shapes.add_picture(item["image"], M, Inches(1.5), width=FULL, height=Inches(4.3))

                    desc = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, M, Inches(6.0), FULL, Inches(1.2))
                    desc.fill.solid()
                    desc.fill.fore_color.rgb = RGBColor(255,255,255)
                    desc.text_frame.text = item["text"]

                binary = io.BytesIO()
                prs.save(binary)
                binary.seek(0)

                st.session_state.generated_ppt_binary = binary
                st.session_state.generated_filename = final_filename
                st.rerun()

    else:
        st.download_button("Download Report", st.session_state.generated_ppt_binary, file_name=st.session_state.generated_filename, use_container_width=True)
        if st.button("Reset / New Report", use_container_width=True):
            st.session_state.report_items = []
            st.session_state.generated_ppt_binary = None
            st.rerun()
