import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import io
from datetime import datetime

# Set page to wide mode
st.set_page_config(page_title="Report Generator", layout="wide")

st.title("Field Inspection Report Generator")
st.markdown("---")

# --- 1. SESSION STATE SETUP ---
if "report_items" not in st.session_state:
    st.session_state.report_items = []
if "edit_index" not in st.session_state:
    st.session_state.edit_index = None
if "generated_ppt_binary" not in st.session_state:
    st.session_state.generated_ppt_binary = None
if "generated_filename" not in st.session_state:
    st.session_state.generated_filename = ""

# --- 2. CALLBACK FUNCTIONS (Fixes the Error) ---
def add_entry_callback():
    """Adds the item and clears inputs BEFORE the page reruns."""
    # Get values from session state directly
    uploaded_file = st.session_state.get("single_uploader")
    description = st.session_state.get("entry_desc")
    category = st.session_state.get("cat_selector")
    custom_cat = st.session_state.get("custom_cat_input")

    # Determine final category
    final_cat = category
    if category == "Other..." and custom_cat:
        final_cat = custom_cat
    
    if uploaded_file and description:
        # Add to list
        st.session_state.report_items.append({
            "category": final_cat,
            "text": description,
            "image": uploaded_file
        })
        
        # Clear the inputs safely
        st.session_state["entry_desc"] = ""
        st.session_state["single_uploader"] = None
        # We don't need st.rerun() here, callbacks automatically trigger a rerun
        
        # Reset generation since data changed
        st.session_state.generated_ppt_binary = None
    else:
        st.error("Please provide both an image and a description.")

def save_edit_callback():
    """Saves changes and exits edit mode."""
    idx = st.session_state.edit_index
    uploaded_file = st.session_state.get("single_uploader")
    description = st.session_state.get("entry_desc")
    category = st.session_state.get("cat_selector")
    custom_cat = st.session_state.get("custom_cat_input")

    final_cat = category
    if category == "Other..." and custom_cat:
        final_cat = custom_cat

    # Use new image if uploaded, otherwise keep old one
    current_item = st.session_state.report_items[idx]
    final_img = uploaded_file if uploaded_file else current_item["image"]

    st.session_state.report_items[idx] = {
        "category": final_cat,
        "text": description,
        "image": final_img
    }
    
    # Exit edit mode and clear inputs
    st.session_state.edit_index = None
    st.session_state["entry_desc"] = ""
    st.session_state["single_uploader"] = None
    st.session_state.generated_ppt_binary = None

def cancel_edit_callback():
    st.session_state.edit_index = None
    st.session_state["entry_desc"] = ""
    st.session_state["single_uploader"] = None

def delete_item_callback(index):
    st.session_state.report_items.pop(index)
    if st.session_state.edit_index == index:
        st.session_state.edit_index = None
    st.session_state.generated_ppt_binary = None

def edit_item_callback(index):
    st.session_state.edit_index = index
    # Pre-fill the description box
    st.session_state["entry_desc"] = st.session_state.report_items[index]["text"]
    st.session_state.generated_ppt_binary = None


# --- 3. SETTINGS SIDEBAR ---
with st.sidebar:
    st.header("Report Settings")
    report_title = st.text_input("Report Title", "Field Inspection Report")
    date_option = st.selectbox("Date Format", ["Month & Year", "Date Only (MM-DD-YYYY)", "Date & Time", "Custom Text"])

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
            selected_time = st.time_input("Select Time", datetime.now())
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
    batch_files = st.file_uploader("Select Multiple Images", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
    
    if st.button("Add All Batch Images", type="primary"):
        if batch_files:
            for file in batch_files:
                st.session_state.report_items.append({
                    "category": "Exterior",
                    "text": "",
                    "image": file
                })
            st.session_state.generated_ppt_binary = None
            st.success(f"Added {len(batch_files)} images! Scroll down to edit.")
        else:
            st.warning("No files selected.")


# --- 5. SINGLE INPUT / EDIT FORM ---
is_editing = st.session_state.edit_index is not None
edit_item = st.session_state.report_items[st.session_state.edit_index] if is_editing else None

with st.container():
    st.markdown("### " + (f"Editing Page {st.session_state.edit_index + 1}" if is_editing else "Add Single Entry"))
    c1, c2 = st.columns([1, 1])

    with c1:
        standard_options = ["Exterior", "Interior"]
        default_ix = 0
        custom_val = ""

        if is_editing:
            if edit_item["category"] in standard_options:
                default_ix = standard_options.index(edit_item["category"])
            else:
                default_ix = 2
                custom_val = edit_item["category"]

        st.selectbox("Category", standard_options + ["Other..."], index=default_ix, key="cat_selector")
        
        # Only show custom input if 'Other...' is selected
        if st.session_state.get("cat_selector") == "Other...":
            st.text_input("Enter Custom Category", value=custom_val, key="custom_cat_input")

        # Description Input
        st.text_area("Description", height=150, placeholder="Enter observation here...", key="entry_desc")

    with c2:
        if is_editing:
            st.image(edit_item["image"], width=150, caption="Current Image")
            st.caption("Leave upload blank to keep current image.")
        
        st.file_uploader("Upload Image (Single)", type=["png", "jpg", "jpeg"], key="single_uploader")

    st.write("")
    b1, b2 = st.columns([1, 6])

    if is_editing:
        b1.button("Save Changes", type="primary", on_click=save_edit_callback)
        b2.button("Cancel Edit", on_click=cancel_edit_callback)
    else:
        # NOTICE: using on_click handles the clearing safely
        st.button("Add Entry", type="primary", on_click=add_entry_callback)


# --- 6. PREVIEW LIST (Fixed Layout) ---
if st.session_state.report_items:
    st.markdown("---")
    st.subheader(f"Current Entries ({len(st.session_state.report_items)})")
    
    for i, item in enumerate(st.session_state.report_items):
        with st.container():
            col_img, col_det, col_act = st.columns([2, 5, 1])
            
            with col_img:
                st.image(item["image"], use_container_width=True)
            
            with col_det:
                st.markdown(f"### Page {i+1}")
                st.markdown(f"**Category:** {item['category']}")
                
                # --- FIX: Single Line Description ---
                if item['text'] == "":
                    st.markdown("**Description:** *No description yet*")
                else:
                    # Using write lets it wrap naturally, or markdown for bold prefix
                    st.markdown(f"**Description:** {item['text']}")
            
            with col_act:
                st.button("Edit", key=f"ed_{i}", on_click=edit_item_callback, args=(i,))
                st.button("Delete", key=f"del_{i}", on_click=delete_item_callback, args=(i,))
            
            st.divider()


# --- 7. PPT GENERATION LOGIC ---
if st.session_state.report_items:
    if st.session_state.generated_ppt_binary is None:
        if st.button("Generate Report", type="primary", use_container_width=True):
            prs = Presentation()

            # Title Slide
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = report_title
            slide.placeholders[1].text = report_subtitle

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

                # Header
                header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, MARGIN_X, TOP_Y, COL_WIDTH, HEADER_HEIGHT)
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

                # Text Box
                desc_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, MARGIN_X, TOP_Y + HEADER_HEIGHT, COL_WIDTH, BODY_HEIGHT)
                desc_box.fill.solid()
                desc_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
                desc_box.line.color.rgb = RGBColor(0, 0, 0)
                tf = desc_box.text_frame
                tf.text = item["text"]
                tf.vertical_anchor = MSO_ANCHOR.TOP
                tf.margin_top = Inches(0.2)
                tf.margin_left = Inches(0.2)
                p = tf.paragraphs[0]
                p.font.bold = False
                p.font.size = Pt(20)
                p.font.color.rgb = RGBColor(0, 0, 0)
                p.alignment = PP_ALIGN.LEFT
                tf.word_wrap = True

                # Image
                img_x = MARGIN_X + COL_WIDTH + GAP
                try:
                    item["image"].seek(0)
                except:
                    pass
                pic = slide.shapes.add_picture(item["image"], img_x, TOP_Y, width=COL_WIDTH, height=IMG_BLOCK_HEIGHT)
                pic.line.color.rgb = RGBColor(0, 0, 0)
                pic.line.width = Pt(1)

                # Footer
                footer_y = SLIDE_HEIGHT - Inches(0.5)
                footer_box = slide.shapes.add_textbox(MARGIN_X, footer_y, Inches(4), Inches(0.5))
                fp = footer_box.text_frame.paragraphs[0]
                fp.text = report_title
                fp.font.size = Pt(10)
                fp.font.color.rgb = RGBColor(80, 80, 80)
                
                page_box = slide.shapes.add_textbox(SLIDE_WIDTH - MARGIN_X - Inches(2), footer_y, Inches(2), Inches(0.5))
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
            st.rerun()
