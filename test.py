cell = 'B22'
while not cell.isdigit():
    cell = cell[1:]
print(cell)