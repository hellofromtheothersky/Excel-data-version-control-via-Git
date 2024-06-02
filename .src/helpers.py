def find_duplicates(lst :list) -> dict[any, list[int]]:
    duplicate_positions = {}
    
    for i, item in enumerate(lst):
        if item in duplicate_positions:
            duplicate_positions[item].append(i)
        else:
            duplicate_positions[item] = [i]
    
    return {key: value for key, value in duplicate_positions.items() if len(value) > 1}