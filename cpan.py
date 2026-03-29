import pandas as pd
import networkx as nx

df = pd.read_csv("links.csv")

data = df[['A End', 'Z End']]

# Clean IPs
data['A End'] = data['A End'].str.strip().str.split('_').str[0]
data['Z End'] = data['Z End'].str.strip().str.split('_').str[0]

# Create graph
G = nx.Graph()
G.add_edges_from(data.values)

source = "10.121.24.167"
target = "10.121.24.130"

if nx.has_path(G, source, target):
    paths = nx.all_simple_paths(G, source, target, cutoff=5)

    for i, path in enumerate(paths, 1):
        print(f"Path {i}: {' -> '.join(path)}")

        if i == 5:   # limit output
            break
else:
    print("No path exists")