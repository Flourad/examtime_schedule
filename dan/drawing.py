#-*- coding: utf-8 -*-
from xlrd import open_workbook
import matplotlib.pyplot as plt
import networkx as nx
import random

data = dict()
docs = {}
color_list  = {}
edge_num = 0

def read_xls(wb):  
    for s in wb.sheets():
        for row in range(2,s.nrows):
            values = []
            for col in range(s.ncols):                            
                values.append(s.cell(row, col).value)
            exam = values[0]
            if exam not in docs:
                docs[exam] = []
            docs[exam].append(' '.join([values[1], values[2]]))
            if exam not in data:
                data[exam] = []
            data[exam].append(('_'.join([values[1], values[2]]), values[3], values[4]))   
        break
    return docs

def build_graph(nodes):
    # created networkx graph 
    global edge_num
    G = nx.Graph()
    G.add_nodes_from(nodes)
    total_edges = [(x, y) for i, x in enumerate(nodes) for y in nodes[i+1:]]
    for node1, node2  in total_edges:
        if list(set(docs[node1]) & set(docs[node2])):
            G.add_edge(node1, node2) 
            edge_num = edge_num + 1       
    return G
              
def coloring(G,nodes): 
    nodes_color = {} 
    for node in nodes:
        nodes_color[node]=0        
    for node in nodes:
        neighbor_node=nx.all_neighbors(G,node)
        neighbor_node_colors= [nodes_color[node1] for node1 in neighbor_node]
        for k in range(1,10000):
            if k not in neighbor_node_colors:
                nodes_color[node]=k 
                break   
            
    color_list = [nodes_color[node] for node in nodes]
    return color_list,nodes_color 

def drawing(G,nodes,color_list):      
    # draw graph
    pos = nx.shell_layout(G)
    nx.draw(G,pos,nodelist=nodes,node_size=500,node_color=color_list,width=0.2,lable='3.xcl')
    #show graph
    plt.show()  
    
def test(docs,nodes_color):
    color_stu = {}
    for color in list(set(nodes_color.values())):
        color_stu[color] = []
    for k,v in nodes_color.items():
        color_stu[v].extend(docs[k])
    for k,v in color_stu.items():
        print "color ",k,":","初始学生人数:",len(v),", ","去除重复后的学生人数:",len(list(set(v)))
        #print "学生： ",','.join(v)
       

if __name__ == "__main__":
    wb1 = open_workbook('c:/isima3.xls')
    docs = read_xls(wb1)
    nodes1 = docs.keys()
    G1 = build_graph(nodes1)
    min_color_num = 10000
    min_nodes = nodes1[:]
    for _ in range(100):
        random.shuffle(nodes1) 
        (color_list,nodes_color)  = coloring(G1,nodes1)
        color_num = len(set(color_list))
        #print color_num
        if color_num < min_color_num:
            min_color_num = color_num
            min_nodes = nodes1[:]
    
    nodes = min_nodes[:]
    
    (color_list,nodes_color)  = coloring(G1,nodes1)
    
    print "The numble of nodes:",len(nodes)
    print "The numble of edges:",edge_num
    print "The kind of color:",min_color_num
    
    test(docs,nodes_color)
    
    drawing(G1,nodes1,color_list)
    
  
   
    
    