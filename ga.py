import os
import streamlit.components.v1 as components

GA_ID = "G-8DM8073S27"

def _is_local() -> bool:
    try:
        headers = st.context.headers
        host = headers.get("Host", "")
        return (
            "localhost" in host or
            "127.0.0.1" in host or
            "0.0.0.0"   in host
        )
    except Exception:
        return False  # if headers not accessible, assume deployed

def load_ga(app_name: str) -> None:
    try:
        if _is_local():
            return  # skip entirely for localhost

        components.html(
            f"""
            <script>
            (function() {{
                const w = window.parent;

                let clientId = null;
                try {{
                    clientId = w.localStorage.getItem("custom_ga_user_id");
                    if (!clientId) {{
                        clientId = "user_" + Math.random().toString(36).slice(2) + "_" + Date.now();
                        w.localStorage.setItem("custom_ga_user_id", clientId);
                    }}
                }} catch(e) {{
                    clientId = "user_" + Math.random().toString(36).slice(2);
                }}

                try {{
                    const existing = w.document.querySelector(
                        'script[src*="googletagmanager.com/gtag/js?id={GA_ID}"]'
                    );
                    if (!existing) {{
                        const s = w.document.createElement("script");
                        s.async = true;
                        s.src = "https://www.googletagmanager.com/gtag/js?id={GA_ID}";
                        w.document.head.appendChild(s);
                    }}
                }} catch(e) {{
                    console.warn("GA script injection failed:", e);
                    return;
                }}

                try {{
                    w.dataLayer = w.dataLayer || [];
                    function gtag() {{ w.dataLayer.push(arguments); }}
                    w.gtag = w.gtag || gtag;
                }} catch(e) {{
                    console.warn("GA dataLayer init failed:", e);
                    return;
                }}

                if (!w.__ga_initialized__) {{
                    w.__ga_initialized__ = true;
                    setTimeout(() => {{
                        try {{
                            w.gtag('js', new Date());
                            w.gtag('config', '{GA_ID}', {{
                                user_id:    clientId,
                                app_name:   '{app_name}',
                                page_title: '{app_name}'
                            }});

                            const loadKey = "ga_loaded_{app_name}";
                            if (!w.sessionStorage.getItem(loadKey)) {{
                                w.gtag('event', 'app_loaded', {{
                                    app_name: '{app_name}',
                                    user_id:  clientId
                                }});
                                w.sessionStorage.setItem(loadKey, "true");
                            }}
                        }} catch(e) {{
                            console.warn("GA config/event failed:", e);
                        }}
                    }}, 800);
                }}
            }})();
            </script>
            """,
            height=0,
        )
    except Exception:
        pass  # never crash the app due to GA


def track_event(event_name: str, **params) -> None:
    try:
        if _is_local():
            return

        import json
        params_json = json.dumps(params)

        components.html(
            f"""
            <script>
            (function() {{
                try {{
                    const w = window.parent;
                    if (!w.gtag) return;

                    let uid = "anonymous";
                    try {{
                        uid = w.localStorage.getItem("real_user_id") 
                           || w.localStorage.getItem("custom_ga_user_id") 
                           || "anonymous";
                    }} catch(e) {{}}

                    const params = {params_json};
                    params.user_id = uid;
                    params.page_path = window.location.pathname;

                    w.gtag('event', '{event_name}', params);

                }} catch(e) {{
                    console.warn("GA track_event failed:", e);
                }}
            }})();
            </script>
            """,
            height=0,
        )
    except Exception:
        pass

# <!-- Google tag (gtag.js) -->
# <script async src="https://www.googletagmanager.com/gtag/js?id=G-M59HEQJCB3"></script>
# <script>
#   window.dataLayer = window.dataLayer || [];
#   function gtag(){dataLayer.push(arguments);}
#   gtag('js', new Date());

#   gtag('config', 'G-M59HEQJCB3');
# </script>