from collections import defaultdict
from gurobipy import *
from pandas import read_excel
filename = "buad5092-m7-final-integration-US-mileage.xlsx"

def opt(sheet, data_type):

  # data
  m = Model()
  data = read_excel(filename, sheet_name=f"{sheet}-{data_type}")
  data_demand = read_excel(filename, sheet_name=f"{sheet}-demand")
  num_rows, num_cols = data.shape[0], data.shape[1] - 1
  locs_rows, locs_cols = [data.iloc[i][0] for i in range(num_rows)], list(data.columns.values)[1:]
  demands = {locs_cols[i].replace(" ", "-"): int(data_demand.get(locs_cols[i])) for i in range(num_cols)}
  # remove spaces in city names for Gurobi format files
  locs_rows, locs_cols = [i.replace(" ", "-") for i in locs_rows], [i.replace(" ", "-") for i in locs_cols]

  # variables
  service_centers, is_serviced_by = defaultdict(list), defaultdict(list)
  is_service_center_vars = {locs_cols[i]: m.addVar(vtype=GRB.BINARY, name=f"is_service_center_{locs_cols[i]}") for i in range(num_cols)}
  for i in range(num_cols):
    for j in range(num_rows):
      v = m.addVar(vtype=GRB.BINARY, name=f"{locs_cols[i]}_is_service_center_for_{locs_rows[j]}")
      service_centers[locs_rows[j]].append(v)
      is_serviced_by[locs_cols[i]].append(v)
  v = None
  m.update()

  # constraints
  m.addLConstr(quicksum(is_service_center_vars.values()), GRB.EQUAL, 3)
  for i in range(num_cols):
    m.addLConstr(quicksum(is_serviced_by[locs_cols[i]]), GRB.LESS_EQUAL, num_rows)
    m.addGenConstrOr(is_service_center_vars[locs_cols[i]], is_serviced_by[locs_cols[i]])
    for j in range(num_rows):  m.addLConstr(quicksum(service_centers[locs_rows[j]]), GRB.EQUAL, 1)
  m.update()

  # objective
  distances = [quicksum([x * y for (x,y) in zip(service_centers[locs_rows[i]], data.loc[i][1:])]) for i in range(num_rows)]
  obj = [(demands[locs_cols[i]] * distances[i] / 1000) for i in range(num_cols)]
  m.setObjective(quicksum(obj), GRB.MINIMIZE)
  m.update()
  m.optimize()
  # m.display()
  print(m.objVal)

  # format output
  s, t = defaultdict(list), defaultdict(list)
  for i in m.getVars():
    if i.x > 0:
      if "is_service_center_for_" in i.varName:
        s[i.varName[:i.varName.index("_")]].append(i.varName[i.varName.rindex("_")+1:])
  for i, j in s.items():  t[i.replace("-", " ")] = [k.replace("-", " ") for k in j]
  tt = [{k: v} for (k, v) in t.items()]

  print("\n\nOUTPUT: \n\n")
  print(f"The minimized total distance is {round(m.objVal, 3)} {data_type}.\n\n")
  print(f"The 3 manufacturing sites are located at {list(t.keys())[0]} and {list(t.keys())[1]} and {list(t.keys())[2]}.\n\n")
  print(f"The manufacturing site assignments are: \n\n  {tt[0]} \n\n  {tt[1]} \n\n  {tt[2]}\n\n")
  print("done.\n\n")
  return (round(m.objVal, 3))

# opt("ex65", "miles")
# opt("st", "miles")
# opt("st", "hours")
