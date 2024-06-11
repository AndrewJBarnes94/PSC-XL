import argparse

def parse_and_print_log(file_path, log_level=None):
    log_levels = ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']
    if log_level and log_level.upper() not in log_levels:
        print(f"Invalid log level: {log_level}. Valid levels are: {', '.join(log_levels)}")
        return
    
    with open(file_path, 'r') as file:
        for line in file:
            stripped_line = line.strip()
            if stripped_line:  # Check if the line is not empty after stripping
                if log_level:
                    if f" - {log_level.upper()} - " in stripped_line:
                        print(stripped_line)
                else:
                    print(stripped_line)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Parse and print log file.")
    parser.add_argument('file_path', help="Path to the log file")
    parser.add_argument('-l', '--log_level', help="Filter logs by level (DEBUG, INFO, WARNING, ERROR, CRITICAL)", default=None)
    
    args = parser.parse_args()
    
    parse_and_print_log(args.file_path, args.log_level)
