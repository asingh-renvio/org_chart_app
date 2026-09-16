# HR organization chart Streamlit app

This app accepts an HR Excel export and rebuilds the single-page organization chart from that upload. Its renderer follows the colours, individual manager team boxes, three-person team rows, department ordering, and alignment approach established in `org_chart_pdf.py`.

## Run on the server

Graphviz must be available on the server as `dot`.

```bash
cd /Users/amitsingh/Documents/transcript_task
source .venv/bin/activate
pip install -r org_chart_streamlit/requirements.txt
streamlit run org_chart_streamlit/app.py --server.address 0.0.0.0 --server.port 8501
```

Run it in `screen` for a persistent service:

```bash
screen -S hr-org-chart
cd /Users/amitsingh/Documents/transcript_task
source .venv/bin/activate
streamlit run org_chart_streamlit/app.py --server.address 0.0.0.0 --server.port 8501
```

Detach with `Ctrl+A`, then `D`. Reattach with `screen -r hr-org-chart`.

## Template policy

The upload is the source of truth for employee names, titles, departments, manager paths, additions, and removals. Known departments retain their approved styling. New department values create new department panels, and new direct managers create new team panels automatically.