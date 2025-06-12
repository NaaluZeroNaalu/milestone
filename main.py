import streamlit as st

pages = {

    "REPORTS":[
        
        st.Page("Milestone.py",title="Milestone",icon=":material/home:"),
        st.Page("MilestoneFinishing.py",title="MilestoneFinishing",icon=":material/home:"),
       

    ]
}


st.navigation(pages).run()
