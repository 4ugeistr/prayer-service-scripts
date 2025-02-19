def stretch_texts(qty_of_verses_requested, texts):
    result = []
    repetitions_per_text, remaining_verses = divmod(qty_of_verses_requested, len(texts))
    for t in texts:
        result.append([t]*repetitions_per_text)
    for i in range(remaining_verses):
        result[i].append(result[i][0])
    return sum(result,[])

# Example usage
qty_of_verses_requested = 8
texts = ["1", "2", "3"]

stretched_text = stretch_texts(qty_of_verses_requested, texts)
print(stretched_text)    # Output: ['1', '1', '1', '2', '2', '2', '3', '3']
