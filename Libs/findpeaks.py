import numpy as np

def PeakFinder(progress_bar,
               yvalues,
               tolerance,
               minPeakDistance = 0,
               minMaximaValue = np.nan,
               maxMaximaValue = np.nan,
               excludeOnEdges = False,
            ):

    xvalues = np.arange(len(yvalues))

    # tolerance = np.std(yvalues)

    maxima = find_maxima(yvalues, tolerance, excludeOnEdges, progress_bar)
    minima = find_minima(yvalues, tolerance, excludeOnEdges, progress_bar)

    if minMaximaValue is not np.nan:
        maxima = trim_peak_height(yvalues, maxima, False)
    if maxMaximaValue is not np.nan:
        minima = trim_peak_height(yvalues, minima, True)
    if minPeakDistance > 0:
        maxima = trim_peak_distance(maxima, xvalues, minPeakDistance)
        minima = trim_peak_distance(minima, xvalues, minPeakDistance)

    return xvalues, yvalues, maxima, minima


def find_maxima(xx, tolerance, edge_mode, progress_bar):
    INCLUDE_EDGE = 0
    CIRCULAR = 2
    len_x = len(xx)
    orig_len = len_x
    if len_x < 2:
        return []
    if tolerance < 0:
        tolerance = 0
    if edge_mode == CIRCULAR:
        cascade3 = np.empty(len_x * 3)
        for jj in range(len_x):
            cascade3[jj] = xx[jj]
            cascade3[jj + len_x] = xx[jj]
            cascade3[jj + 2*len_x] = xx[jj]
        len_x *= 3
        xx = cascade3
    max_positions = np.zeros(len_x, dtype=int)
    max_val = xx[0]
    min_val = xx[0]
    max_pos = 0
    last_max_pos = -1
    left_valley_found = (edge_mode == INCLUDE_EDGE)
    max_count = 0
    for jj in range(1, len_x):
        progress_bar['value'] = (jj + 1) / len_x * 100
        progress_bar.update()
        val = xx[jj]
        if val > min_val + tolerance:
            left_valley_found = True
        if val > max_val and left_valley_found:
            max_val = val
            max_pos = jj
        if left_valley_found:
            last_max_pos = max_pos
        if val < max_val - tolerance and left_valley_found:
            max_positions[max_count] = max_pos
            max_count += 1
            left_valley_found = False
            min_val = val
            max_val = val
        if val < min_val:
            min_val = val
            if not left_valley_found:
                max_val = val
    if edge_mode == INCLUDE_EDGE:
        if max_count > 0 and max_positions[max_count - 1] != last_max_pos:
            max_positions[max_count] = last_max_pos
            max_count += 1
        elif max_count == 0 and max_val - min_val >= tolerance:
            max_positions[max_count] = last_max_pos
            max_count += 1
    cropped = max_positions[:max_count]
    max_positions = cropped
    max_values = np.empty(max_count)
    for jj in range(max_count):
        pos = max_positions[jj]
        mid_pos = pos
        while pos < len_x - 1 and xx[pos] == xx[pos + 1]:
            mid_pos += 0.5
            pos += 1
        max_positions[jj] = int(mid_pos)
        max_values[jj] = xx[max_positions[jj]]
    rank_positions = np.argsort(max_values)
    return_arr = np.empty(max_count, dtype=int)
    for jj in range(max_count):
        pos = max_positions[rank_positions[jj]]
        return_arr[max_count - jj - 1] = pos  # use descending order
    if edge_mode == CIRCULAR:
        count = 0
        for jj in range(len(return_arr)):
            pos = return_arr[jj] - orig_len
            if 0 <= pos < orig_len:  # pick maxima from cascade center part
                return_arr[count] = pos
                count += 1
        return_arr = return_arr[:count]
    return return_arr

def find_minima(xx, tolerance, edge_mode, progress_bar):
    len_x = len(xx)
    neg_arr = [-x for x in xx]
    min_positions = find_maxima(neg_arr, tolerance, edge_mode, progress_bar)
    return min_positions


def trim_peak_height(positions, minima, yvalues):
    size1 = len(positions)
    size2 = 0
    for i in range(size1):
        if filtered_height(yvalues[positions[i]], minima):
            size2 += 1
        else:
            break  # positions are sorted by amplitude
    newpositions = [positions[i] for i in range(size2)]
    return newpositions

def filtered_height(height, minima, max_minima_value, min_maxima_value):
    if minima:
        return height < max_minima_value
    else:
        return height > min_maxima_value

def trim_peak_distance(positions, xvalues, min_peak_distance):
    size = len(positions)
    temp = [0 for _ in range(size)]
    newsize = 0
    for i in range(size - 1, -1, -1):
        pos1 = positions[i]
        trim = False
        for j in range(i - 1, -1, -1):
            pos2 = positions[j]
            if abs(xvalues[pos2] - xvalues[pos1]) < min_peak_distance:
                trim = True
                break
        if not trim:
            temp[newsize] = pos1
            newsize += 1
    newpositions = [temp[i] for i in range(newsize)]
    return newpositions