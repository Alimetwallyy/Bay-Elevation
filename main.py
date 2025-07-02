import streamlit as st
import io
import string
from pptx import Presentation
from pptx.util import Inches, Cm, Pt
from pptx.enum.shapes import MSO_SHAPE

# --- Main PowerPoint Generation Function ---
def generate_elevation_powerpoint(bay_types_data):
    """
    Creates a PowerPoint presentation from bay type data.
    Each bay type gets its own slide with a technical drawing.
    """
    prs = Presentation()
    # Set slide size to widescreen 16:9
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_slide_layout = prs.slide_layouts[6] 
    
    for bay_type in bay_types_data:
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # Add a title to the slide
        title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12.33), Inches(0.5))
        title_text = title_shape.text_frame
        p = title_text.paragraphs[0]
        p.text = bay_type['name']
        p.font.bold = True
        p.font.size = Pt(24)

        # --- Drawing Logic ---
        # Constants for drawing position and scale
        start_x = Inches(2.0)
        start_y = Inches(7.0) 
        scale = Cm(0.2) # This scales the cm measurements to fit the slide
        current_y = start_y
        
        # Draw shelves from the bottom of the slide to the top
        for i, shelf_name in enumerate(reversed(bay_type['shelves'])):
            shelf_info = bay_type['shelf_details'][shelf_name]
            num_bins = shelf_info['num_bins']
            bin_h = shelf_info['h'] * scale
            bin_w = shelf_info['w'] * scale
            
            shelf_height = bin_h
            shelf_width = num_bins * bin_w

            # Draw the main shelf rectangle (acts as a container)
            slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, start_x, current_y - shelf_height, shelf_width, shelf_height)
            
            # Draw the individual bin dividers within the shelf
            for j in range(num_bins):
                bin_x = start_x + (j * bin_w)
                slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, bin_x, current_y - shelf_height, bin_w, shelf_height)

            # Add the shelf letter label (e.g., "A", "B") to the left of the shelf
            label_box = slide.shapes.add_textbox(start_x - Inches(0.6), current_y - shelf_height, Inches(0.5), shelf_height)
            label_box.text_frame.text = shelf_name
            label_box.text_frame.paragraphs[0].font.size = Pt(10)

            # Add dimension annotations to the right of the shelf
            # This logic only adds the text if the dimensions are different from the shelf below it
            is_first_of_dim = True
            if i > 0:
                prev_shelf_name = list(reversed(bay_type['shelves']))[i-1]
                if bay_type['shelf_details'][prev_shelf_name] == shelf_info:
                    is_first_of_dim = False
            
            if is_first_of_dim:
                dim_x = start_x + shelf_width + Inches(0.2)
                dim_y = current_y - (shelf_height / 2)
                dim_text = f"H: {shelf_info['h']}cm\nW: {shelf_info['w']}cm\nD: {shelf_info['d']}cm"
                dim_box = slide.shapes.add_textbox(dim_x, dim_y - Inches(0.25), Inches(1.5), Inches(0.5))
                tf = dim_box.text_frame
                tf.text = dim_text
                for para in tf.paragraphs:
                    para.font.size = Pt(9)

            # Move the drawing cursor up for the next shelf, adding a small gap
            current_y -= (shelf_height + Cm(0.2))

    # Save the final presentation to a memory buffer
    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer

# --- Streamlit App UI ---
st.set_page_config(layout="wide")
st.title("Bay Elevation Generator 📐")
st.markdown("Define your bay types, their shelves, and bin configurations to generate a PowerPoint elevation drawing.")
st.info("ℹ️ **Note:** This application requires `python-pptx`. Please ensure it is in your requirements.txt file.")

num_bay_types = st.number_input("How many bay types do you want to define?", min_value=1, max_value=50, value=1, key="num_bay_types")

bay_types_data = []

for i in range(num_bay_types):
    # Use session_state to keep track of bay type names dynamically
    if f"bay_type_name_{i}" not in st.session_state:
        st.session_state[f"bay_type_name_{i}"] = f"Bay Type {i + 1}"

    def update_bay_type_name(idx=i):
        st.session_state[f"bay_type_name_{idx}"] = st.session_state[f"bay_type_name_input_{idx}"]

    header = st.session_state[f"bay_type_name_{i}"].strip() or f"Bay Type {i + 1}"

    with st.expander(header, expanded=True):
        bay_name = st.text_input("Bay Type Name", value=st.session_state[f"bay_type_name_{i}"], key=f"bay_type_name_input_{i}", on_change=update_bay_type_name, args=(i,))
        
        shelf_count = st.number_input("Number of shelves in this bay?", min_value=1, max_value=26, value=3, key=f"elevation_shelf_count_{i}")
        shelf_names = list(string.ascii_uppercase[:shelf_count])
        
        shelf_details = {}
        for shelf_name in shelf_names:
            st.divider()
            st.markdown(f"**Configuration for Shelf {shelf_name}**")
            
            num_bins = st.number_input("Number of bins in this shelf?", min_value=1, max_value=50, value=5, key=f"num_bins_{i}_{shelf_name}")
            
            st.markdown(f"**Bin Dimensions for all bins in Shelf {shelf_name} (cm)**")
            c1, c2, c3 = st.columns(3)
            with c1:
                bin_h = st.number_input("Height", min_value=1.0, value=10.0, key=f"bin_h_{i}_{shelf_name}")
            with c2:
                bin_w = st.number_input("Width", min_value=1.0, value=10.0, key=f"bin_w_{i}_{shelf_name}")
            with c3:
                bin_d = st.number_input("Depth", min_value=1.0, value=10.0, key=f"bin_d_{i}_{shelf_name}")
            
            shelf_details[shelf_name] = {'num_bins': num_bins, 'h': bin_h, 'w': bin_w, 'd': bin_d}

        # Add the collected data to our list if a name has been provided
        if bay_name.strip():
            bay_types_data.append({
                "name": bay_name.strip(),
                "shelves": shelf_names,
                "shelf_details": shelf_details
            })

# --- Generate Button and Download Logic ---
if st.button("Generate PowerPoint Elevation", key="generate_ppt"):
    # Basic validation to ensure bay types have names
    if any(not bay['name'] for bay in bay_types_data):
        st.error("Please provide a name for all bay types.")
    else:
        with st.spinner("Generating PowerPoint file..."):
            try:
                ppt_buffer = generate_elevation_powerpoint(bay_types_data)
                
                st.success(f"✅ Success! Generated a PowerPoint with {len(bay_types_data)} slides.")
                st.download_button(
                    label="📥 Download PowerPoint File",
                    data=ppt_buffer,
                    file_name="bay_elevations.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"An error occurred during PowerPoint generation: {e}")