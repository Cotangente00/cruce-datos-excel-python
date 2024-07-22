sequence = 'AAAABBBCCDAABBB'

def unique_in_order(sequence):
    unique_sequence = []
    for caracter in sequence:
        if caracter == sequence[caracter]:
            unique_sequence.append(caracter)
    return unique_sequence

unique_in_order(sequence) 