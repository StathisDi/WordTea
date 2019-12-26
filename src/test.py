import WordTea_classes
from WordTea_classes import reference_list
print(reference_list.__doc__)
print(reference_list.__init__.__doc__)
parent_cl = reference_list("parent", "tag", "label", 1)
child = reference_list("child", "tag", "label", 1, parent_cl)
parent_cl.build_list("text parent", 1)
child.build_list("text child", 1)
parent_cl.build_list("text parent 2", 1)

print("Child:")
print(child.ref_list)
print(child.counter)
print(child.parent_count)
print("Parent")
print(child.parent.ref_list)
print(child.parent.counter)
print(child.parent.parent_count)