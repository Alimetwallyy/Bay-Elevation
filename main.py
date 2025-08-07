import streamlit as st
import plotly.graph_objects as go
import json
import os
from datetime import datetime
import pandas as pd

# Initialize session state for layout data
if 'layout_data' not in st.session_state:
    st.session_state.layout_data = []

st.title("Advanced Warehouse Layout Designer")

# Sidebar for navigation
st.sidebar.header("Options")
action = st.sidebar.radio("Action", ["Design New Layout", "Load Saved Layout"])

if action == "Design New Layout":
    # Input for number of MODs/Areas
    num_mods = st.number_input("Number of MODs/Areas", min_value=1, value=2, step=1)

    # Collect AISLE details
    aisle_data = []
    for i in range(num_mods):
        st.subheader(f"MOD {i+1}")
        with st.expander(f"Configure MOD {i+1}"):
            num_aisles = st.number_input(f"Number of AISLEs in MOD {i+1}", min_value=1, value=2, step=1)
            for j in range(num_aisles):
                col1, col2 = st.columns(2)
                with col1:
                    purpose = st.selectbox(f"AISLE {j+1} Purpose (MOD {i+1})", ["General", "Groceries", "Electronics", "Clothing"], key=f"purpose_{i}_{j}")
                with col2:
                    number = st.number_input(f"AISLE {j+1} Number (MOD {i+1})", min_value=1, value=j+1, step=1, key=f"number_{i}_{j}")
                aisle_data.append({"MOD": i+1, "AISLE Number": number, "Purpose": purpose})

    # Save to session state
    if st.button("Save Layout"):
        st.session_state.layout_data = aisle_data
        st.success("Layout saved to session!")

    # Visualize
    if aisle_data or st.session_state.layout_data:
        data_to_visualize = aisle_data if not st.session_state.layout_data else st.session_state.layout_data
        df = pd.DataFrame(data_to_visualize)

        # Plotly visualization
        fig = go.Figure()
        colors = {"General": "green", "Groceries": "pink", "Electronics": "blue", "Clothing": "orange"}
        for mod in df['MOD'].unique():
            mod_data = df[df['MOD'] == mod]
            for _, row in mod_data.iterrows():
                fig.add_trace(go.Bar(
                    y=[f"MOD {mod} - AISLE {row['AISLE Number']}"],
                    x=[1],
                    orientation='h',
                    marker_color=colors[row['Purpose']],
                    text=row['Purpose'],
                    textposition='auto'
                ))

        fig.update_layout(
            title="Warehouse Layout",
            xaxis_title="",
            yaxis_title="",
            bargap=0.2
        )
        st.plotly_chart(fig)

        # Summary
        st.write("### Summary")
        st.table(df)

        # Legend
        st.write("### Legend")
        for purpose, color in colors.items():
            st.write(f"- {purpose}: <span style='color:{color}'>█</span>", unsafe_allow_html=True)

        # Save to file (JSON and SVG)
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Save to JSON"):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"layout_{timestamp}.json"
                with open(filename, 'w') as f:
                    json.dump(data_to_visualize, f)
                st.success(f"Layout saved as {filename}")
        with col2:
            if st.button("Export to SVG"):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                svg_filename = f"layout_{timestamp}.svg"
                fig.write_image(svg_filename, format="svg")
                with open(svg_filename, "rb") as f:
                    st.download_button(
                        label="Download SVG",
                        data=f,
                        file_name=svg_filename,
                        mime="image/svg+xml"
                    )
                    os.remove(svg_filename)  # Clean up temporary file
                st.success(f"SVG exported as {svg_filename}")

elif action == "Load Saved Layout":
    uploaded_file = st.file_uploader("Upload JSON Layout File", type="json")
    if uploaded_file:
        layout_data = json.load(uploaded_file)
        st.session_state.layout_data = layout_data
        df = pd.DataFrame(layout_data)
        st.write("### Loaded Layout")
        st.table(df)

        # Visualize loaded layout
        fig = go.Figure()
        colors = {"General": "green", "Groceries": "pink", "Electronics": "blue", "Clothing": "orange"}
        for mod in df['MOD'].unique():
            mod_data = df[df['MOD'] == mod]
            for _, row in mod_data.iterrows():
                fig.add_trace(go.Bar(
                    y=[f"MOD {mod} - AISLE {row['AISLE Number']}"],
                    x=[1],
                    orientation='h',
                    marker_color=colors[row['Purpose']],
                    text=row['Purpose'],
                    textposition='auto'
                ))

        fig.update_layout(
            title="Loaded Warehouse Layout",
            xaxis_title="",
            yaxis_title="",
            bargap=0.2
        )
        st.plotly_chart(fig)
