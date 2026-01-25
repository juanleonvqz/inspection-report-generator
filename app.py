import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import io
from datetime import datetime

# --- CONFIG & STYLING ---
st.set_page_config(page_title="Report Generator", layout="wide")

# Inject Custom CSS to fix text visibility issues in inputs and cards
st.markdown("""
<style>
    /* Force input text color to be visible */
    .stTextInput input, .stTextArea textarea {
        color: #333333 !important;
    }
    /* Style for the 'Card' look */
    .report-card {
        background-color: #f8f9fa;
        border: 1px solid #ddd;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 10px;
    }
    .report-label {
        font-size: 0.75rem;
        font-weight: 700;
        color: #555;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 4px;
        display: block;
    }
    .report-value {
        font-size: 1rem;
        color: #000000 !important; /* Force Black Text */
        font-family: sans-serif;
        line-height: 1.5;
    }
    .highlight-warning {
        background-color: #fff3cd;
        color: #856404;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

st.title("Field Inspection Report Generator")
st.markdown("---")

# --- 1. SESSION STATE SETUP ---
if "report_items" not in st.session_state:
    st.session_state.report_items = []
if "edit_index" not in st.session_state:
    st.session_state.edit_index = None # Stores the index of the item currently being edited
if "generated_ppt_binary" not in st.session_state:
    st.session_state.generated_ppt_binary = None
if "uploader_key_id" not in st.session_state:
    st.session_state.uploader_key_id = 0

# --- 2. FUNCTIONS ---

def add_new_entry():
    """Callback to add a single new item from the form."""
    img = st.session_state.get(f"new_img_{st.session_state.uploader_key_id}")
    desc = st.session_state.get("new_desc")
    cat_select = st.session_state.get("new_cat_select")
    cat_custom = st.session_state.get("new_cat_custom")

    final_cat = cat_custom if cat_select == "Other..." else cat_select

    if img and desc:
        st.session_state.report_items.append({
            "category": final_cat,
            "text": desc,
            "image": img
        })
        # Reset Inputs
        st.session_state["new_desc"] = ""
        st.session_state["new_cat_custom"] = ""
        st.session_state.uploader_key_id += 1 # Forces new uploader widget
        st.session_state.generated_ppt_binary = None
    else:
        st.error("Please add both an Image and a Description.")

def add_batch_images():
    """Callback to process batch upload."""
    files = st.session_state.get("batch_uploader")
    if files:
        for f in files:
            st.session_state.report_items.append({
                "category": "Exterior", # Default
                "text": "", # Default empty
                "image": f
            })
        st.session_state.generated_ppt_binary = None
        # Note: We don't need to clear the batch uploader manually as it persists
        # but the logic won't duplicate unless button is pressed again.

def save_inline_edit(index):
    """Callback to save changes made in the inline edit form."""
    new_cat_select = st.session_state.get(f"edit_cat_select_{index}")
    new_cat_custom = st.session_state.get(f"edit_cat_custom_{index}")
    new_desc = st.session_state.get(f"edit_desc_{index}")
    new_img = st.session_state.get(f"edit_img_{index}")
    new_position = st.session_state.get(f"edit_pos_{index}")

    final_cat = new_cat_custom if new_cat_select == "Other..." else new_cat_select
    
    current_item = st.session_state.report_items[index]
    final_image = new_img if new_img is not None else current_item['image']

    updated_item = {
        "category": final_cat,
        "text": new_desc,
        "image": final_image
    }

    # Handle Reordering
    if new_position is not None:
        target_index = new_position - 1
        # Ensure target is within bounds
        target_index = max(0, min(target_index, len(st.session_state.report_items) - 1))
        
        if target_index != index:
            st.session_state.report_items.pop(index)
            st.session_state.report_items.insert(target_index, updated_item)
        else:
            st.session_state.report_items[index] = updated_item
    else:
        st.session_state.report_items[index] = updated_item

    st.session_state.edit_index = None
    st.session_state.generated_ppt_binary = None

def delete_item(index):
    st.session_state.report_items.pop(index)
    if st.session_state.edit_index == index:
        st.session_state.edit_index = None
    st.session_state.generated_ppt_binary = None

def enter_edit_mode(index):
    st.session_state.edit_index = index

# --- 3. SETTINGS SIDEBAR ---
with st.sidebar:
    st.header("Report Settings")
    report_title = st.text_input("Report Title", "Field Inspection Report")
    report_subtitle = st.text_input("Subtitle / Date", datetime.now().strftime("%B %Y"))
    final_filename = f"{report_title.replace(' ', '_')}.pptx"


# --- 4. BATCH UPLOAD (Restored) ---
with st.expander("📂 Batch Upload (Add Multiple Images)", expanded=False):
    st.write("Select multiple images to add them all at once. They will be added with default 'Exterior' category.")
    st.file_uploader("Select Images", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True, key="batch_uploader")
    st.button("Add Batch Images", on_click=add_batch_images)


# --- 5. SINGLE ADD FORM ---
st.subheader("Add Single Slide")
with st.container():
    c1, c2, c3 = st.columns([1, 2, 2])
    
    with c1:
        cat_options = ["Exterior", "Interior", "Roof", "Plumbing", "Electrical", "Other..."]
        st.selectbox("Category", cat_options, key="new_cat_select")
        if st.session_state.get("new_cat_select") == "Other...":
            st.text_input("Custom Category", key="new_cat_custom")
    
    with c2:
        st.text_area("Description", height=100, key="new_desc", placeholder="Type observation here...")
    
    with c3:
        st.file_uploader("Image", type=['jpg','png','jpeg'], key=f"new_img_{st.session_state.uploader_key_id}")
        st.write("")
        st.button("➕ Add Slide", type="primary", on_click=add_new_entry, use_container_width=True)

# --- 6. MAIN LIST (Inline Editing & Display) ---
if st.session_state.report_items:
    st.markdown("---")
    st.subheader(f"Slides ({len(st.session_state.report_items)})")
    
    for i, item in enumerate(st.session_state.report_items):
        
        # === EDIT MODE ===
        if st.session_state.edit_index == i:
            with st.container():
                st.info(f"Editing Slide {i+1}")
                ec1, ec2 = st.columns([1, 1])
                
                with ec1:
                    # Page Reordering
                    st.number_input("Page Order", min_value=1, max_value=len(st.session_state.report_items), value=i+1, key=f"edit_pos_{i}")
                    
                    # Category Logic
                    edit_cat_opts = ["Exterior", "Interior", "Roof", "Plumbing", "Electrical", "Other..."]
                    curr_cat = item['category']
                    try:
                        sel_index = edit_cat_opts.index(curr_cat)
                    except ValueError:
                        sel_index = 5 # 'Other...'
                    
                    st.selectbox("Category", edit_cat_opts, index=sel_index, key=f"edit_cat_select_{i}")
                    
                    # Show custom input if needed
                    if st.session_state.get(f"edit_cat_select_{i}") == "Other...":
                         # Pre-fill with current category if it's not in the standard list
                        val = curr_cat if curr_cat not in edit_cat_opts else ""
                        st.text_input("Custom Category", value=val, key=f"edit_cat_custom_{i}")

                    st.text_area("Description", value=item['text'], height=150, key=f"edit_desc_{i}")

                with ec2:
                    st.image(item['image'], width=200, caption="Current Image")
                    st.file_uploader("Replace Image (Optional)", type=['jpg','png'], key=f"edit_img_{i}")

                btn_c1, btn_c2 = st.columns([1, 4])
                with btn_c1:
                    st.button("💾 Save", key=f"save_{i}", type="primary", on_click=save_inline_edit, args=(i,))
                with btn_c2:
                    st.button("Cancel", key=f"cancel_{i}", on_click=lambda: st.session_state.update(edit_index=None))
                st.markdown("---")

        # === VIEW MODE ===
        else:
            with st.container():
                col_img, col_info, col_act = st.columns([2, 4, 1])
                
                with col_img:
                    st.image(item['image'], use_container_width=True)
                
                with col_info:
                    st.markdown(f"### Slide {i+1}")
                    
                    # Category Card
                    st.markdown(f"""
                        <div class="report-card">
                            <span class="report-label">Category</span>
                            <div class="report-value">{item['category']}</div>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    # Description Card
                    desc_html = item['text'] if item['text'] else '<span class="highlight-warning">NO DESCRIPTION</span>'
                    st.markdown(f"""
                        <div class="report-card">
                            <span class="report-label">Description</span>
                            <div class="report-value">{desc_html}</div>
                        </div>
                    """, unsafe_allow_html=True)

                with col_act:
                    st.button("✏️ Edit", key=f"btn_edit_{i}", on_click=enter_edit_mode, args=(i,), use_container_width=True)
                    st.button("🗑️ Delete", key=f"btn_del_{i}", on_click=delete_item, args=(i,), use_container_width=True)
                
                st.divider()

# --- 7. GENERATE PPT ---
if st.session_state.report_items:
    st.write("")
    if st.button("Generate PowerPoint Report", type="primary", use_container_width=True):
        
        prs = Presentation()
        
        # Title Slide
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = report_title
        slide.placeholders[1].text = report_subtitle
        
        # Layout Config
        MARGIN = Inches(0.5)
        TOP = Inches(1.0)
        WIDTH = Inches(4.25)
        HEIGHT = Inches(5.0)
        
        for idx, item in enumerate(st.session_state.report_items):
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank
            
            # Header
            header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(0.8))
            header.fill.solid()
            header.fill.fore_color.rgb = RGBColor(50, 50, 50)
            header.text = f"{item['category']} - Slide {idx+1}"
            header.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            header.text_frame.paragraphs[0].font.size = Pt(24)
            
            # Text Box
            tb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, MARGIN, TOP, WIDTH, HEIGHT)
            tb.fill.background()
            tb.line.color.rgb = RGBColor(0,0,0)
            tf = tb.text_frame
            tf.text = item['text']
            tf.paragraphs[0].font.size = Pt(18)
            tf.paragraphs[0].font.color.rgb = RGBColor(0,0,0)
            tf.word_wrap = True
            
            # Image
            try: item['image'].seek(0)
            except: pass
            
            pic = slide.shapes.add_picture(item['image'], Inches(5.0), TOP, width=WIDTH, height=HEIGHT)
            pic.line.color.rgb = RGBColor(0,0,0)
            pic.line.width = Pt(1)

        # Save
        binary_output = io.BytesIO()
        prs.save(binary_output)
        binary_output.seek(0)
        
        st.download_button(
            label="⬇️ Download .pptx",
            data=binary_output,
            file_name=final_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="secondary"
        )
