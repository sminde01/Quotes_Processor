# NOTE: DO NOT RUN THIS FILE AS A STREAMLIT SERVER
# NOTE: RUN IT AS `python gtag.py`

# ----- required modules
# streamlit
import streamlit as st

# html parser
from bs4 import BeautifulSoup

# utils
from shutil import copy as shcopy

# misc
from pathlib import Path

def inject_ga_tag() -> None:
    """Function to inject Google Analytics script in the head section of the index html file.

    Params
    ------
    None

    Returns
    -------
    None
    """

    gtag_js = """
    <!-- Global site tag (gtag.js) - Google Analytics -->
    <script async src="https://www.googletagmanager.com/gtag/js?id=G-8XZPRE65W4"></script>
    <script>
        window.dataLayer = window.dataLayer || [];
        function gtag(){dataLayer.push(arguments);}
        gtag('js', new Date());
        gtag('config', 'G-8XZPRE65W4');
    </script>
    <div id="G-8XZPRE65W4"></div>
    """

    gtag_id = "G-8XZPRE65W4"

    # step 1: identify html path of streamlit's index
    index_path = Path(st.__file__).parent / "static" / "index.html"
    print(f"Editing: {index_path}")

    # step 2: initiate html parser
    soup = BeautifulSoup(index_path.read_text(), features = "html.parser")

    # step 3: inject the google analytics script only if already not present
    if not soup.find(id = gtag_id):
        bck_index = index_path.with_suffix(".bck")

        # backup recovery
        if bck_index.exists():
            shcopy(bck_index, index_path)
        
        # save backup
        else:
            shcopy(index_path, bck_index)
        
        # get html
        html = str(soup)

        # inject script
        new_html = html.replace("<head>", "<head>\n" + gtag_js)
        
        # update streamlit's index
        index_path.write_text(new_html)

inject_ga_tag()