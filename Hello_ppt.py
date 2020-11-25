import networkx as nx
from network2ppt.convert import convert

G = nx.DiGraph()

# node ID sd title text
G.add_edge(0, 1)
G.add_edge(4, 0)

# add content
nx.set_node_attributes(G, 
    {
        0: 'Nothing', 
        1: '-------', 
        4: 'xxxxxxxxx',
    },
    'content'
)
# change title color
G.nodes[0]['color'] = (100, 80, 170)
# change content color
G.nodes[1]['bg'] = (180, 50, 100)


convert(G, 'hello_world.pptx', 
       scale=[1.5, 0.6],  # scale ratio of [x, y], shape of slide is [1, 1]
       connector_type=2)