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

# -- Helper to clear inputs --
def clear_form_inputs():
    """Resets specific session state widgets to blank."""
    if "entry_desc" in st.session_state:
        st.session_state["entry_desc"] = "" 
    if "single_uploader" in st.session_state:
        st.session_state["single_uploader"] = None
    # Optional: Reset category to default if desired
    # if "cat_selector" in st.session_state:
    #     st.session_state["cat_selector"] = "Exterior"

def clear_edit_mode():
    st.session_state.edit_index = None
    clear_form_inputs()

def reset_generation():
    """Forces the 'Download' button to revert to 'Generate' if data changes."""
    st.session_state.generated_ppt_binary = None

# --- 2. SETTINGS SIDEBAR ---
with st.sidebar:
    st.header("Report Settings")

    report_title = st.text_input("Report Title", "Field Inspection Report")

    date_option = st.selectbox(
        "Date Format",
        ["Month & Year", "Date Only (MM-DD-YYYY)", "Date & Time", "Custom Text"],
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


# --- 3. BATCH UPLOAD SECTION ---
with st.expander("Batch Upload (Add Multiple Images)", expanded=False):
    st.write("Select all images in your folder and drag them here to add them all at once.")
    st.write("Works best on Computer.")
    batch_files = st.file_uploader("Select Multiple Images", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
    
    if st.button("Add All Batch Images", type="primary"):
        if batch_files:
            count = 0
            for file in batch_files:
                st.session_state.report_items.append({
                    "category": "Exterior", # Default category
                    "text": "", # Empty description
                    "image": file
                })
                count += 1
            
            reset_generation()
            st.success(f"Successfully added {count} images! Scroll down to edit them.")
        else:
            st.warning("No files selected.")


# --- 4. SINGLE INPUT / EDIT FORM ---
is_editing = st.session_state.edit_index is not None
edit_item = (
    st.session_state.report_items[st.session_state.edit_index] if is_editing else None
)

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

        options_final = standard_options + ["Other..."]
        cat_select = st.selectbox("Category", options_final, index=default_ix, key="cat_selector")

        category_final = cat_select
        if cat_select == "Other...":
            category_final = st.text_input("Enter Custom Category", value=custom_val)

        # LOGIC CHANGE: If editing, pre-fill the session state key so the text area shows it
        if is_editing and "entry_desc" not in st.session_state:
             st.session_state["entry_desc"] = edit_item["text"]
        
        # We use a key "entry_desc" so we can programmatically clear it later
        description = st.text_area("Description", height=150, placeholder="Enter observation here...", key="entry_desc")

    with c2:
        if is_editing:
            st.image(edit_item["image"], width=150, caption="Current Image")
            st.caption("To keep current image, leave upload blank.")
        
        # Key is crucial here to clear it later
        uploaded_file = st.file_uploader("Upload Image (Single)", type=["png", "jpg", "jpeg"], key="single_uploader")

    st.write("")
    b1, b2 = st.columns([1, 6])

    if is_editing:
        if b1.button("Save Changes", type="primary"):
            final_img = uploaded_file if uploaded_file else edit_item["image"]
            
            st.session_state.report_items[st.session_state.edit_index] = {
                "category": category_final,
                "text": description,
                "image": final_img,
            }
            clear_edit_mode() # This also clears the inputs
            reset_generation()
            st.rerun()
            
        if b2.button("Cancel Edit"):
            clear_edit_mode()
            st.rerun()
    else:
        if st.button("Add Entry", type="primary"):
            if uploaded_file and description:
                st.session_state.report_items.append(
                    {
                        "category": category_final,
                        "text": description,
                        "image": uploaded_file,
                    }
                )
                st.success("Entry added.")
                
                # --- RESET LOGIC ---
                clear_form_inputs() # Wipes the text area and uploader
                reset_generation()
                st.rerun() # Refresh page to show blank inputs
            else:
                st.error("Please provide both an image and a description.")


# --- 5. PREVIEW LIST (UPDATED LAYOUT) ---
if st.session_state.report_items:
    st.markdown("---")
    st.subheader(f"Current Entries ({len(st.session_state.report_items)})")
    
    for i, item in enumerate(st.session_state.report_items):
        with st.container():
            # NEW LAYOUT: Image | Details | Actions
            col_img, col_det, col_act = st.columns([2, 5, 1])
            
            with col_img:
                st.image(item["image"], use_container_width=True)
            
            with col_det:
                st.markdown(f"### Page {i+1}")
                st.markdown(f"**Category:** {item['category']}")
                st.markdown("**Description:**")
                if item['text'] == "":
                    st.warning("No description yet")
                else:
                    st.text(item["text"]) # Using st.text for cleaner block look
            
            with col_act:
                if st.button("Edit", key=f"ed_{i}"):
                    st.session_state.edit_index = i
                    # Pre-load the description into the input box key for the next rerun
                    st.session_state["entry_desc"] = item["text"]
                    reset_generation()
                    st.rerun()
                
                if st.button("Delete", key=f"del_{i}"):
                    st.session_state.report_items.pop(i)
                    if st.session_state.edit_index == i:
                        clear_edit_mode()
                    reset_generation()
                    st.rerun()
            
            st.divider()


# --- 6. PPT GENERATION LOGIC ---
if st.session_state.report_items:

    # CHECK: Do we already have a generated file?
    if st.session_state.generated_ppt_binary is None:

        if st.button("Generate Report", type="primary", use_container_width=True):
            
            # --- GENERATION START ---
            prs = Presentation()

            # Title Slide
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = report_title
            slide.placeholders[1].text = report_subtitle

            # Dimensions & Config
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

                # --- LEFT COLUMN ---
                # 1. Header
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

                # 2. Description Box
                desc_box = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    MARGIN_X,
                    TOP_Y + HEADER_HEIGHT,
                    COL_WIDTH,
                    BODY_HEIGHT,
                )
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

                # --- RIGHT COLUMN (FORCE FIT IMAGE) ---
                img_x = MARGIN_X + COL_WIDTH + GAP
                img_y = TOP_Y

                # Important: Reset file pointer for batch images
                try:
                    item["image"].seek(0)
                except:
                    pass

                pic = slide.shapes.add_picture(
                    item["image"],
                    img_x,
                    img_y,
                    width=COL_WIDTH,
                    height=IMG_BLOCK_HEIGHT,
                )
                pic.line.color.rgb = RGBColor(0, 0, 0)
                pic.line.width = Pt(1)

                # --- FOOTER ---
                footer_y = SLIDE_HEIGHT - Inches(0.5)

                # 1. Report Name
                footer_box = slide.shapes.add_textbox(
                    MARGIN_X, footer_y, Inches(4), Inches(0.5)
                )
                fp = footer_box.text_frame.paragraphs[0]
                fp.text = report_title
                fp.font.size = Pt(10)
                fp.font.color.rgb = RGBColor(80, 80, 80)

                # 2. Page Number
                page_box = slide.shapes.add_textbox(
                    SLIDE_WIDTH - MARGIN_X - Inches(2), footer_y, Inches(2), Inches(0.5)
                )
                pp = page_box.text_frame.paragraphs[0]
                pp.text = f"Page {index + 1}"
                pp.font.size = Pt(10)
                pp.font.color.rgb = RGBColor(80, 80, 80)
                pp.alignment = PP_ALIGN.RIGHT

            # Save
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
            reset_generation()
            st.rerun()
