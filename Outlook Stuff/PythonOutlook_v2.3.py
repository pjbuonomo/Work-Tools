def parse_line(line):
    size_first_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+)\s+(bid|offered|offer)\s*(?:@|at|-)?\s*(\d*\.\d+)?"
    name_first_pattern = r"([\w\s-]+?)\s+\((\w+)\)\s+(\d*\.\d+)?\s*(bid|offered|offer)\s*(?:@|at|-)?\s*(\d+\.\d+)"
    dual_action_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+)\s+(bid)\s+/\s+(\d+\.\d+)\s+(offer)"

    default_dict = {"Name": "", "Size": "", "CUSIP": "", "Actions": "", "Price": "", "Error": line}

    entries = []

    # Size-First Format
    if re.match(size_first_pattern, line):
        for match in re.finditer(size_first_pattern, line):
            size, name, cusip, price, action, alt_price = match.groups()[0], match.groups()[3], match.groups()[4], match.groups()[5], match.groups()[6], match.groups()[7]
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": action, "Price": price if price else alt_price, "Error": ""})

    # Name-First Format
    elif re.match(name_first_pattern, line):
        for match in re.finditer(name_first_pattern, line):
            name, cusip, alt_price, action, price = match.groups()[0], match.groups()[1], match.groups()[2], match.groups()[3], match.groups()[4]
            entries.append({"Name": name.strip(), "Size": "", "CUSIP": cusip, "Actions": action, "Price": price if price else alt_price, "Error": ""})

    # Dual-Action Format
    elif re.match(dual_action_pattern, line):
        for match in re.finditer(dual_action_pattern, line):
            size, name, cusip, bid_price, offer_price = match.groups()[0], match.groups()[3], match.groups()[4], match.groups()[5], match.groups()[7]
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "bid", "Price": bid_price, "Error": ""})
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "offer", "Price": offer_price, "Error": ""})

    return entries if entries else [default_dict]

Please show in all offerings. Many thanks for the focus.
Please show in all offerings. Many thanks for the focus.




























2.25mm Gateway 2023-3 A (36779CAF3) offered @ 107.10




Please show in all offerings. Many thanks for the focus.

























Tailwind 2022-1 B (87403TAE6) bid at 99









500k Res Re 2020-I 13 (76124AAB4) offered @ **BH trades and we care to buy more**

Axed for more...
4mm Tailwind 2022-1 B (87403TAE6) offered @ **BH trades**
2.25mm Gateway 2023-3 A (36779CAF3) offered @ 107.00
Please show in all offerings. Many thanks for the focus. Many bids higher.























Tailwind 2022-1 B (87403TAE6) bid at 99








2.25mm Gateway 2023-3 A (36779CAF3) 102 bid / 107.00 offer



74.50 bid for 3264 Re 2022-1 (88577CAB7)
64.50 bid for Herbie 2021-1 A (42703VAE3)
Please let us know if you care to offer... Thanks
Axed to buy more...
3mm Gateway 2022-1 A (36779CAA4) offered @ **BH trades**
3mm Mystic 2021-2 B (62865LAC1) - 97.35 bid / 98.10 offer
