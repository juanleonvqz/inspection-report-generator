import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import io

st.title("📋 Report Generator")
st.write("Upload photos and descriptions to generate a formatted report.")

# Initialize session state to store the list of report items
if 'report_items' not in st.session_state:
    st.session_state.report_items = []

# --- INPUT SECTION ---
with st.container():
    st.subheader("New Entry")
    
    col1, col2 = st.columns(2)
    with col1:
        # 1. The Category Input (Interior vs Exterior)
        category = st.radio("Category", ["Exterior", "Interior"], horizontal=True)
        # 2. The Description
        description = st.text_area("Observation / Description", height=150)
    
    with col2:
        # 3. The Image
        uploaded_image = st.file_uploader("Upload Photo", type=['png', 'jpg', 'jpeg'])

    if st.button("Add Entry"):
        if uploaded_image and description:
            st.session_state.report_items.append({
                "category": category,
                "text": description,
                "image": uploaded_image
            })
            st.success(f"Added {category} item!")
        else:
            st.error("Please provide both an image and a description.")

# --- PREVIEW SECTION ---
if len(st.session_state.report_items) > 0:
    st.divider()
    st.write(f"**Current Items: {len(st.session_state.report_items)}**")
    for i, item in enumerate(st.session_state.report_items):
        st.text(f"{i+1}. [{item['category']}] {item['text'][:50]}...")

# --- GENERATION LOGIC ---
if st.button("Generate Report PPT"):
    if len(st.session_state.report_items) == 0:
        st.warning("Add some items first!")
    else:
        # Create blank presentation
        prs = Presentation()
        
        # We use a BLANK layout (index 6) so we can draw our own boxes
        # Standard slides are 10 inches wide x 7.5 inches tall
        
        for item in st.session_state.report_items:
            slide = prs.slides.add_slide(prs.slide_layouts[6]) 

            # --- 1. DRAW HEADER BLOCK ---
            # Position: Top of page
            header_left = Inches(0.5)
            header_top = Inches(0.5)
            header_width = Inches(9)   # Spans across page
            header_height = Inches(0.5)
            
            # Add the shape (Rectangle)
            header_shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, header_left, header_top, header_width, header_height
            )
            
            # Style the Header
            header_shape.text = item['category']
            header_shape.fill.solid()
            header_shape.fill.fore_color.rgb = RGBColor(200, 200, 200) # Light Grey
            header_shape.line.color.rgb = RGBColor(0, 0, 0) # Black border
            
            # Format Text inside Header
            paragraph = header_shape.text_frame.paragraphs[0]
            paragraph.font.bold = True
            paragraph.font.size = Pt(18)
            paragraph.font.color.rgb = RGBColor(0, 0, 0)

            # --- 2. DRAW TEXT BLOCK (LEFT) ---
            text_left = Inches(0.5)
            text_top = Inches(1.1) # Below header
            text_width = Inches(4.4)
            text_height = Inches(4.5)
            
            # Create a text box with a border (using a Rectangle shape with no fill)
            text_box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, text_left, text_top, text_width, text_height
            )
            text_box.fill.background() # No fill (transparent)
            text_box.line.color.rgb = RGBColor(0, 0, 0) # Black border
            
            # Add the user's text
            text_frame = text_box.text_frame
            text_frame.text = item['text']
            text_frame.margin_top = Inches(0.1)
            text_frame.margin_left = Inches(0.1)
            
            # --- 3. DRAW IMAGE BLOCK (RIGHT) ---
            img_left = Inches(5.1) # To the right of the text block
            img_top = Inches(1.1)
            img_width = Inches(4.4)
            img_height = Inches(4.5)
            
            # Add the picture
            # Note: We set both width and height to force it to "Fit the Block"
            # This ensures perfect alignment with the text box.
            pic = slide.shapes.add_picture(
                item['image'], img_left, img_top, width=img_width, height=img_height
            )
            
            # Add a border to the picture to match the style
            line = pic.line
            line.color.rgb = RGBColor(0, 0, 0)
            line.width = Pt(1)

        # Output file
        binary_output = io.BytesIO()
        prs.save(binary_output)
        binary_output.seek(0)

        st.download_button(
            label="⬇️ Download Report",
            data=binary_output,
            file_name="inspection_report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
