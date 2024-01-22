def parse_line(line):
    # Existing patterns
    size_first_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+)\s+(bid|offered|offer)\s*(?:@|at|-)?\s*(\d*\.\d+)?"
    name_first_pattern = r"([\w\s-]+?)\s+\((\w+)\)\s+(\d*\.\d+)?\s*(bid|offered|offer)\s*(?:@|at|-)?\s*(\d+\.\d+)"
    dual_action_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+)\s+(bid)\s+/\s+(\d+\.\d+)\s+(offer)"
    
    # New pattern for special phrases like "**BH trades**"
    special_phrase_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(offered|bid)\s*(?:@|at|-)?\s*(\*\*.*?\*\*)"

    default_dict = {"Name": "", "Size": "", "CUSIP": "", "Actions": "", "Price": "", "Error": line}

    entries = []

    # Check each pattern
    if re.match(size_first_pattern, line) or re.match(name_first_pattern, line) or re.match(dual_action_pattern, line):
        for pattern in [size_first_pattern, name_first_pattern, dual_action_pattern]:
            for match in re.finditer(pattern, line):
                size, name, cusip, action, price = '', '', '', '', ''
                if pattern == size_first_pattern:
                    size, name, cusip, price, action, alt_price = match.groups()
                elif pattern == name_first_pattern:
                    name, cusip, alt_price, action, price = match.groups()
                elif pattern == dual_action_pattern:
                    size, name, cusip, bid_price, offer_price = match.groups()
                    entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "offer", "Price": offer_price, "Error": ""})
                
                price = price if price else alt_price
                entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": action, "Price": price, "Error": ""})
                break

    # Handle special phrase pattern
    elif re.match(special_phrase_pattern, line):
        for match in re.finditer(special_phrase_pattern, line):
            size, name, cusip, action, special_phrase = match.groups()
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": action, "Price": special_phrase, "Error": ""})

    return entries if entries else [default_dict]

