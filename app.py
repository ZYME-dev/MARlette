import streamlit as st

pg = st.navigation(
    {
        "": [
            st.Page("pages/marlette.py", title="MARlette ⭐", default=True,),
            
        ],
        "Aide": [
            # st.Page("pages/why_marlette.py", title="Why MARlette ⭐?!", icon="🤷"),
            st.Page("pages/manual_user.py", title="Manuel - basique", icon="📗"),
            st.Page("pages/manual_dev.py", title="Manuel - avancé", icon="📕"),
        ],
    },
    expanded=True
)

pg.run()
