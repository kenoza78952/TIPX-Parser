# debug_utils.py

LOG_BUFFER = [] 

def debug(message, level="INFO"):
    line = f"[{level}] {message}"
    LOG_BUFFER.append(line)
    print(line)

def save_debug_log(filepath):
    with open(filepath, "w", encoding="utf-8") as f:
        for line in LOG_BUFFER:
            f.write(line + "\n")
