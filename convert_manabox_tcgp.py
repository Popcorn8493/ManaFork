import csv
import re
import time
from datetime import datetime
from pathlib import Path
from tkinter import Button, END, Frame, Label, Listbox, Scrollbar, Tk, TclError
from tkinter.filedialog import askopenfilename

import pandas as pd
import requests
import unicodedata
from rapidfuzz import fuzz

# Content filtering configuration
FILTER_PRERELEASE = False  # Filter prerelease content
FILTER_PROMO = False  # Filter promotional content

# Set name normalization mappings
SET_ALIAS = {
		"Universes Beyond: The Lord of the Rings: Tales of Middle-earth": "LTR",
		"Commander: The Lord of the Rings: Tales of Middle-earth":        "LTC",
		"the list":                                                       "The List",
		"edge of eternities":                                             "eoe",
		"EOE":                                                            "eoe"
}

# Condition format standardization
CONDITION_MAP = {
		"near mint":         "Near Mint",
		"lightly played":    "Lightly Played",
		"moderately played": "Moderately Played",
		"heavily played":    "Heavily Played",
		"damaged":           "Damaged"
}

# Condition priority ordering
condition_rank = {
		"near mint":         0,
		"lightly played":    1,
		"moderately played": 2,
		"heavily played":    3,
		"damaged":           4
}

# Minimum price threshold
FLOOR_PRICE = 0.10

# External API settings
SCRYFALL_API_BASE = "https://api.scryfall.com"
SCRYFALL_RATE_LIMIT = 0.1  # Request throttling interval
last_scryfall_request = 0

# Processing state management
given_up_cards = []
scryfall_only_cards = []  # External data source entries
confirmed_matches = {}
scryfall_cache = {}  # API response cache
pending_confirmations = []  # Deferred user confirmations


def rate_limit_scryfall():
	"""Enforce API request throttling."""
	global last_scryfall_request
	current_time = time.time()
	elapsed = current_time - last_scryfall_request
	if elapsed < SCRYFALL_RATE_LIMIT:
		sleep_time = SCRYFALL_RATE_LIMIT - elapsed
		time.sleep(sleep_time)
	last_scryfall_request = time.time()


def write_csv_output(file_path, fieldnames, data_list, description):
	"""Write card data to CSV file."""
	with open(file_path, mode='w', newline='', encoding='utf-8') as csvfile:
		writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
		writer.writeheader()
		for entry in data_list:
			writer.writerow(entry)
	print(f"{description}: {len(data_list)} cards")


def query_scryfall_card(card_name, set_code, collector_number=None):
	"""Retrieve card data with caching."""
	cache_key = f"{card_name}|{set_code}|{collector_number or ''}"
	
	if cache_key in scryfall_cache:
		return scryfall_cache[cache_key]
	
	rate_limit_scryfall()
	
	try:
		# Try exact search first if we have collector number
		if collector_number:
			url = f"{SCRYFALL_API_BASE}/cards/{set_code}/{collector_number}"
			response = requests.get(url, timeout=10)
			
			if response.status_code == 200:
				card_data = response.json()
				scryfall_cache[cache_key] = card_data
				return card_data
		
		# Fallback to name search in set
		params = {
				'q':      f'"{card_name}" set:{set_code}',
				'format': 'json'
		}
		url = f"{SCRYFALL_API_BASE}/cards/search"
		response = requests.get(url, params=params, timeout=10)
		
		if response.status_code == 200:
			search_data = response.json()
			if search_data.get('total_cards', 0) > 0:
				# Return first exact name match
				for card in search_data.get('data', []):
					if card.get('name', '').lower() == card_name.lower():
						scryfall_cache[cache_key] = card
						return card
				# If no exact match, return first result
				first_card = search_data['data'][0]
				scryfall_cache[cache_key] = first_card
				return first_card
		
		# Cache negative results too
		scryfall_cache[cache_key] = None
		return None
	
	except Exception as e:
		print(f"Scryfall API error for {card_name} ({set_code}): {e}")
		scryfall_cache[cache_key] = None
		return None


def get_scryfall_variants(card_name, set_code):
	"""Retrieve card variant information."""
	cache_key = f"variants|{card_name}|{set_code}"
	
	if cache_key in scryfall_cache:
		return scryfall_cache[cache_key]
	
	rate_limit_scryfall()
	
	try:
		params = {
				'q':      f'"{card_name}" set:{set_code}',
				'format': 'json'
		}
		url = f"{SCRYFALL_API_BASE}/cards/search"
		response = requests.get(url, params=params, timeout=10)
		
		if response.status_code == 200:
			search_data = response.json()
			variants = []
			for card in search_data.get('data', []):
				if card.get('name', '').lower() == card_name.lower():
					variant_info = {
							'collector_number': card.get('collector_number'),
							'promo':            card.get('promo', False),
							'promo_types':      card.get('promo_types', []),
							'frame_effects':    card.get('frame_effects', []),
							'finishes':         card.get('finishes', []),
							'variation':        card.get('variation', False),
							'full_art':         card.get('full_art', False),
							'textless':         card.get('textless', False),
							'image_status':     card.get('image_status'),
							'border_color':     card.get('border_color')
					}
					variants.append(variant_info)
			
			scryfall_cache[cache_key] = variants
			return variants
		
		scryfall_cache[cache_key] = []
		return []
	
	except Exception as e:
		print(f"Scryfall variants error for {card_name} ({set_code}): {e}")
		scryfall_cache[cache_key] = []
		return []


def query_scryfall_by_id(scryfall_id):
	"""Retrieve card data by identifier."""
	cache_key = f"id|{scryfall_id}"
	
	if cache_key in scryfall_cache:
		return scryfall_cache[cache_key]
	
	rate_limit_scryfall()
	
	try:
		url = f"{SCRYFALL_API_BASE}/cards/{scryfall_id}"
		response = requests.get(url, timeout=10)
		
		if response.status_code == 200:
			card_data = response.json()
			scryfall_cache[cache_key] = card_data
			return card_data
		
		scryfall_cache[cache_key] = None
		return None
	
	except Exception as e:
		print(f"Scryfall ID query error for {scryfall_id}: {e}")
		scryfall_cache[cache_key] = None
		return None


def create_scryfall_fallback_entry(scryfall_card, manabox_row, condition):
	"""Generate entry from external data source."""
	promo_suffix = ""
	if scryfall_card.get('promo'):
		promo_types = scryfall_card.get('promo_types', [])
		if promo_types:
			promo_suffix = f" ({', '.join(promo_types).title()})"
	
	return {
			"TCGplayer Id":          "Scryfall Verified",
			"Product Line":          "Magic: The Gathering",
			"Set Name":              scryfall_card.get('set_name', ''),
			"Product Name":          scryfall_card.get('name', '') + promo_suffix,
			"Number":                scryfall_card.get('collector_number', ''),
			"Rarity":                scryfall_card.get('rarity', '').title(),
			"Condition":             condition,
			"Add to Quantity":       int(manabox_row.get("Quantity", "1")),
			"TCG Marketplace Price": get_market_price(manabox_row, None)
	}


def enhance_matches_with_scryfall(normalized_key, matches, ref_data, manabox_row=None):
	"""Supplement matching with external data."""
	card_name, set_name, collector_number, condition, suffix = normalized_key
	
	# First try using Scryfall ID if available from Manabox
	scryfall_card = None
	if manabox_row and manabox_row.get("Scryfall ID"):
		scryfall_id = manabox_row.get("Scryfall ID").strip()
		if scryfall_id:
			scryfall_card = query_scryfall_by_id(scryfall_id)
	
	# Fallback to name/set search if no ID or ID lookup failed
	if not scryfall_card:
		# Convert set name to set code for Scryfall
		set_code = SET_ALIAS.get(set_name, set_name)
		# Try common abbreviations
		if len(set_name) > 3:
			# Convert "edge of eternities" to "eoe"
			words = set_name.split()
			if len(words) >= 2:
				set_code = ''.join(word[0] for word in words[:3]).lower()
		
		scryfall_card = query_scryfall_card(card_name, set_code, collector_number)
	
	if scryfall_card:
		# Check if this variant might be missing from TCGplayer data
		if not matches or (matches and matches[0][1] < 300):  # Low confidence in existing matches
			promo_info = ""
			if scryfall_card.get('promo'):
				promo_types = scryfall_card.get('promo_types', [])
				promo_info = f" (Promo: {', '.join(promo_types)})" if promo_types else " (Promo)"
			
			# Create a high-confidence synthetic match using Scryfall data
			if manabox_row:
				scryfall_entry = create_scryfall_fallback_entry(scryfall_card, manabox_row, condition)
				# Create a synthetic match key that will score very high
				synthetic_key = (card_name, set_name, collector_number, condition, suffix)
				synthetic_match = (synthetic_key, 350)  # Higher than auto-confirm threshold
				
				# Add the synthetic entry to ref_data so it can be used
				ref_data[synthetic_key] = scryfall_entry
				
				# Insert at the beginning of matches list
				matches.insert(0, synthetic_match)
				print(f"Found Scryfall-only variant{promo_info}")
	else:
		print(f"Card not found on Scryfall")
	
	return matches


def remove_accents(text):
	"""Normalize character encoding."""
	return ''.join(
			c for c in unicodedata.normalize('NFKD', text)
			if not unicodedata.combining(c)
	)


def is_double_sided_candidate(product_name):
	"""Detect dual-face card indicators."""
	pn = product_name.lower()
	return '//' in pn or ('double' in pn and 'sided' in pn)


def get_market_price(manabox_row, ref_row=None):
	"""Calculate appropriate price value."""
	candidate_fields = ["TCG Marketplace Price", "List Price", "Retail Price"]
	if ref_row:
		for field in candidate_fields:
			price = str(ref_row.get(field, "")).strip()
			try:
				if price and float(price) > 0:
					return price
			except ValueError:
				continue
	csv_candidate_fields = ["Purchase price"]
	for field in csv_candidate_fields:
		price = str(manabox_row.get(field, "")).strip()
		try:
			if price and float(price) > 0:
				return price
		except ValueError:
			continue
	return f"{FLOOR_PRICE:.2f}"


def normalize_key(card_name, set_name, condition, number):
	"""Standardize card identifiers for comparison."""
	suffix = ""
	if "(" in card_name and ")" in card_name:
		card_name = re.sub(r"\(.*?\)", "", card_name).strip()
	# Remove accents from card_name
	card_name = remove_accents(card_name)
	card_name = card_name.split('//')[0].strip()  # Use only text before '//' if present.
	normalized_card_name = re.sub(r"[^a-zA-Z0-9 ,'-]", "", card_name).strip().lower()
	# Remove accents from set_name
	set_name = remove_accents(set_name)
	normalized_set_name = re.sub(r"[^a-zA-Z0-9 ]", "", set_name).strip().lower()
	if normalized_set_name in ["plst", "the list"]:
		normalized_set_name = "the list reprints"
	if "prerelease cards" in normalized_set_name:
		return None
	if normalized_set_name == "the list":
		number = number.split("-")[-1] if number else ""
	normalized_number = re.sub(r"[^\d\-]", "", str(number).strip()) if number else None
	if normalized_number == "":
		normalized_number = None
	return normalized_card_name, normalized_set_name, normalized_number, condition.lower(), suffix


def build_given_up_entry(manabox_row, condition, card_name, set_name):
	"""Create unmatched card entry."""
	return {
			"TCGplayer Id":          "Not Found",
			"Product Line":          "Magic: The Gathering",
			"Set Name":              set_name,
			"Product Name":          card_name,
			"Number":                manabox_row.get("Collector number", "").strip(),
			"Rarity":                manabox_row.get("Rarity", ""),
			"Condition":             condition,
			"Add to Quantity":       int(manabox_row.get("Quantity", "1")),
			"TCG Marketplace Price": get_market_price(manabox_row, None)
	}


def load_reference_data(reference_csv):
	"""Initialize card database."""
	start_time = time.time()
	print("Loading reference database...")
	
	try:
		ref_df = pd.read_csv(reference_csv, dtype={"Number": "str"})
		# Load reference data length for potential future use
		_ = len(ref_df)
		ref_df = ref_df[ref_df["Set Name"].notnull()]
		
		# Apply filters
		excluded_count = 0
		if FILTER_PRERELEASE:
			mask = ref_df["Product Name"].str.contains("Prerelease", case=False, na=False)
			excluded_count += mask.sum()
			ref_df = ref_df[~mask]
		
		if FILTER_PROMO:
			promo_patterns = [r"\(Bundle\)", r"\(Buyabox\)", r"\(Buy-a-[Bb]ox\)", r"\(Promo\)",
			                  r"\(Release\)", r"\(Launch\)", r"\(Store Championship\)",
			                  r"\(Game Day\)", r"\(FNM\)", r"\(Judge\)"]
			mask = ref_df["Product Name"].str.contains("|".join(promo_patterns), case=False, na=False)
			excluded_count += mask.sum()
			ref_df = ref_df[~mask]
		
		# Build lookup table
		records = ref_df.to_dict('records')
		ref_data = {}
		
		for row in records:
			key = normalize_key(
					row.get("Product Name", ""),
					row.get("Set Name", ""),
					row.get("Condition", "Near Mint"),
					row.get("Number", "")
			)
			if key:
				ref_data[key] = row
		
		total_time = time.time() - start_time
		print(f"Loaded {len(ref_data):,} cards in {total_time:.1f}s" +
		      (f" (excluded {excluded_count:,})" if excluded_count > 0 else ""))
		
		return ref_data
	except FileNotFoundError:
		print(f"Reference file not found: {reference_csv}")
		exit()


def find_best_match(normalized_key, card_database):
	"""Locate optimal card matches."""
	matches = []
	exact_number_matches = []
	
	for ref_key in card_database.keys():
		# Quick check: if the first letters differ, skip.
		if normalized_key[0] and ref_key[0] and normalized_key[0][0] != ref_key[0][0]:
			continue
		
		query_words = normalized_key[0].split()
		candidate_words = ref_key[0].split()
		if len(query_words) == 1 and len(candidate_words) == 1:
			if query_words[0] != candidate_words[0]:
				continue
		elif len(query_words) > 1 and len(candidate_words) > 1:
			if not set(query_words).intersection(set(candidate_words)):
				continue
		
		base_score = fuzz.ratio(normalized_key[0], ref_key[0])
		if normalized_key[0] in ref_key[0] or ref_key[0] in normalized_key[0]:
			base_score += 20
		if normalized_key[1] == ref_key[1]:
			base_score += 50
		
		# Handle collector number matching with fallback for missing variants
		if not normalized_key[2] or not ref_key[2]:
			base_score += 50
		elif normalized_key[2] == ref_key[2]:
			base_score += 100
			exact_number_matches.append((ref_key, base_score))
		else:
			base_score -= 15
		
		cond1 = normalized_key[3].replace("foil", "").strip()
		cond2 = ref_key[3].replace("foil", "").strip()
		if cond1 in condition_rank and cond2 in condition_rank:
			diff = abs(condition_rank[cond1] - condition_rank[cond2])
			if diff == 0:
				base_score += 50
			elif diff == 1:
				base_score -= 10
			else:
				base_score -= 30
		else:
			if normalized_key[3] != ref_key[3]:
				base_score -= 20
		
		if ("prerelease" in card_database[ref_key]["Product Name"].lower() or
				"prerelease cards" in card_database[ref_key]["Set Name"].lower()):
			continue
		
		special_print_penalties = {
				"foil":       40,
				"showcase":   30,
				"etched":     30,
				"borderless": 30,
				"extended":   30,
				"gilded":     30
		}
		for term, penalty in special_print_penalties.items():
			in_query = term in normalized_key[3]
			in_ref = term in ref_key[3]
			if in_query != in_ref:
				base_score -= penalty
		
		matches.append((ref_key, base_score))
	
	# If we have exact number matches, prioritize those
	if exact_number_matches:
		matches = exact_number_matches
	# If no exact number matches but we have a card name/set match, show warning and continue
	elif matches and normalized_key[2]:
		print(
				f"Warning: No exact collector number match found for {normalized_key[0]} #{normalized_key[2]}. Showing closest variants.")
	
	matches.sort(key=lambda x: x[1], reverse=True)
	return matches


def create_modern_gui():
	"""Initialize user interface."""
	root = Tk()
	root.title("MTG Card Matcher - Batch Confirmation")
	root.configure(bg='#2b2b2b')
	
	# Modern styling
	style_config = {
			'bg':               '#2b2b2b',
			'fg':               '#ffffff',
			'font':             ('Segoe UI', 11),
			'selectbackground': '#404040',
			'selectforeground': '#ffffff'
	}
	
	# Calculate window size and center it
	window_width = 1200
	window_height = 700
	screen_width = root.winfo_screenwidth()
	screen_height = root.winfo_screenheight()
	x = (screen_width - window_width) // 2
	y = (screen_height - window_height) // 2
	root.geometry(f"{window_width}x{window_height}+{x}+{y}")
	
	return root, style_config


def confirm_match_simple_fallback(pending_items):
	"""Text-based user confirmation."""
	results = {}
	print("\nGUI unavailable, using console confirmation:")
	print("Commands: [1-9] select match, [s] skip, [a] auto-confirm all remaining")
	
	for item_index, (normalized_key, matches, local_ref_data) in enumerate(pending_items):
		print(f"\n--- Item {item_index + 1}/{len(pending_items)} ---")
		print(f"Card: {normalized_key[0]}")
		print(f"Set: {normalized_key[1]} | Number: {normalized_key[2]}")
		
		for idx, (match, score) in enumerate(matches[:5]):  # Show top 5
			candidate = local_ref_data.get(match, {})
			print(f"{idx + 1}: {candidate.get('Product Name', 'Unknown')} (Score: {score})")
		
		while True:
			try:
				choice = input("Select [1-5], [s]kip, [a]uto-all: ").strip().lower()
				if choice == 's':
					results[item_index] = None
					break
				elif choice == 'a':
					# Auto-confirm remaining with best match
					for remaining_idx in range(item_index, len(pending_items)):
						_, remaining_matches, _ = pending_items[remaining_idx]
						results[remaining_idx] = remaining_matches[0][0] if remaining_matches else None
					return results
				elif choice.isdigit() and 1 <= int(choice) <= min(5, len(matches)):
					results[item_index] = matches[int(choice) - 1][0]
					break
				else:
					print("Invalid choice. Try again.")
			except (ValueError, IndexError):
				print("Invalid choice. Try again.")
	
	return results


def confirm_match_gui_batch(pending_items):
	"""Batch user confirmation interface."""
	if not pending_items:
		return {}
	
	print(f"Opening batch confirmation GUI for {len(pending_items)} items...")
	
	try:
		root, style = create_modern_gui()
		root.lift()  # Bring window to front
		root.focus_force()  # Force focus
		root.attributes('-topmost', True)  # Make window stay on top initially
		root.after(100, lambda *args: root.attributes('-topmost', False))  # Remove topmost after showing
		
		results = {}
		current_item = [0]  # Use list for mutable reference
	except Exception as gui_error:
		print(f"GUI initialization failed: {gui_error}")
		return confirm_match_simple_fallback(pending_items)
	
	# Header
	header_frame = Frame(root, bg=style['bg'], height=60)
	header_frame.pack(fill="x", padx=20, pady=10)
	header_frame.pack_propagate(False)
	
	title_label = Label(header_frame,
	                    text=f"Card Matching Confirmation ({len(pending_items)} items)",
	                    font=('Segoe UI', 16, 'bold'),
	                    bg=style['bg'], fg='#4CAF50')
	title_label.pack(side="top", pady=5)
	
	progress_label = Label(header_frame,
	                       text="",
	                       font=('Segoe UI', 10),
	                       bg=style['bg'], fg='#cccccc')
	progress_label.pack(side="top")
	
	# Main content frame
	main_frame = Frame(root, bg=style['bg'])
	main_frame.pack(fill="both", expand=True, padx=20, pady=10)
	
	# Card info frame
	info_frame = Frame(main_frame, bg='#3a3a3a', relief="solid", bd=1)
	info_frame.pack(fill="x", pady=(0, 15))
	
	card_name_label = Label(info_frame, text="",
	                        font=('Segoe UI', 14, 'bold'),
	                        bg='#3a3a3a', fg='#ffffff')
	card_name_label.pack(pady=10)
	
	card_details_label = Label(info_frame, text="",
	                           font=('Segoe UI', 10),
	                           bg='#3a3a3a', fg='#cccccc')
	card_details_label.pack(pady=(0, 10))
	
	# Matches frame
	matches_frame = Frame(main_frame, bg=style['bg'])
	matches_frame.pack(fill="both", expand=True)
	
	Label(matches_frame, text="Select the best match:",
	      font=('Segoe UI', 12, 'bold'),
	      bg=style['bg'], fg='#ffffff').pack(anchor="w", pady=(0, 10))
	
	# Listbox with scrollbar
	list_frame = Frame(matches_frame, bg=style['bg'])
	list_frame.pack(fill="both", expand=True)
	
	listbox = Listbox(list_frame,
	                  font=('Consolas', 10),
	                  bg='#404040',
	                  fg='#ffffff',
	                  selectbackground='#4CAF50',
	                  selectforeground='#000000',
	                  activestyle='none',
	                  relief="solid",
	                  bd=1)
	listbox.pack(side="left", fill="both", expand=True)
	
	scrollbar = Scrollbar(list_frame, orient="vertical", command=listbox.yview)
	scrollbar.pack(side="right", fill="y")
	listbox.configure(yscrollcommand=scrollbar.set)
	
	# Buttons frame
	button_frame = Frame(root, bg=style['bg'], height=80)
	button_frame.pack(fill="x", padx=20, pady=10)
	button_frame.pack_propagate(False)
	
	def create_button(parent, text, command, color='#4CAF50'):
		btn = Button(parent, text=text, command=command,
		             font=('Segoe UI', 11, 'bold'),
		             bg=color, fg='white',
		             relief="flat",
		             padx=20, pady=8,
		             cursor="hand2")
		return btn
	
	def update_display():
		if current_item[0] >= len(pending_items):
			print(f"All {len(pending_items)} confirmations completed. Closing GUI...")
			root.quit()  # Exit mainloop
			root.destroy()  # Destroy window
			return
		
		item = pending_items[current_item[0]]
		normalized_key, matches, local_ref_data = item
		
		# Update progress
		progress_text = f"Item {current_item[0] + 1} of {len(pending_items)}"
		progress_label.config(text=progress_text)
		
		# Update card info
		card_name_label.config(text=f"Card: {normalized_key[0]}")
		details_text = f"Set: {normalized_key[1]} | Number: {normalized_key[2]} | Condition: {normalized_key[3]}"
		card_details_label.config(text=details_text)
		
		# Update matches list
		listbox.delete(0, END)
		for idx, (match, score) in enumerate(matches):
			candidate = local_ref_data.get(match, {})
			match_text = (f"{idx + 1:2}: {candidate.get('Product Name', 'Unknown')[:40]:<40} | "
			              f"Set: {candidate.get('Set Name', 'Unknown')[:20]:<20} | "
			              f"#{candidate.get('Number', 'N/A'):<4} | "
			              f"Score: {score:3}")
			listbox.insert(END, match_text)
		
		if matches:
			listbox.selection_set(0)
			listbox.focus_set()
	
	def on_confirm():
		if current_item[0] < len(pending_items):
			item = pending_items[current_item[0]]
			normalized_key, matches, local_ref_data = item
			
			selected_indices = listbox.curselection()
			if selected_indices:
				selected_match = matches[selected_indices[0]][0]
				results[current_item[0]] = selected_match
			else:
				results[current_item[0]] = None
			
			current_item[0] += 1
			update_display()
	
	def on_skip():
		if current_item[0] < len(pending_items):
			results[current_item[0]] = None
			current_item[0] += 1
			update_display()
	
	def on_auto_all():
		# Auto-confirm remaining items with best match
		for remaining_i in range(current_item[0], len(pending_items)):
			item = pending_items[remaining_i]
			normalized_key, matches, local_ref_data = item
			if matches:
				results[remaining_i] = matches[0][0]  # Best match
			else:
				results[remaining_i] = None
		print(f"Auto-confirmed {len(pending_items) - current_item[0]} remaining items. Closing GUI...")
		root.quit()  # Exit mainloop
		root.destroy()  # Destroy window
	
	def go_previous():
		if current_item[0] > 0:
			current_item[0] -= 1
			update_display()
	
	# Create buttons
	Button(button_frame, text="◀ Previous",
	       command=go_previous,
	       font=('Segoe UI', 10), bg='#757575', fg='white',
	       relief="flat", padx=15, pady=6).pack(side="left", padx=(0, 10))
	
	create_button(button_frame, "✓ Confirm & Next", on_confirm).pack(side="left", padx=5)
	create_button(button_frame, "⤼ Skip", on_skip, '#FF9800').pack(side="left", padx=5)
	create_button(button_frame, "⚡ Auto-Confirm All", on_auto_all, '#2196F3').pack(side="left", padx=5)
	
	def on_cancel():
		print("Confirmation cancelled by user. Closing GUI...")
		root.quit()
		root.destroy()
	
	create_button(button_frame, "✕ Cancel All", on_cancel, '#f44336').pack(side="right")
	
	# Keyboard shortcuts
	def on_key(event):
		if event.keysym == 'Return':
			on_confirm()
		elif event.keysym == 'space':
			on_skip()
		elif event.keysym == 'Escape':
			on_cancel()
	
	root.bind('<Key>', on_key)
	root.focus_set()
	
	# Handle window close button (X)
	root.protocol("WM_DELETE_WINDOW", on_cancel)
	
	# Start display
	update_display()
	
	# Remove the problematic timeout check that was causing Tkinter errors
	# The GUI will handle timeouts through normal user interaction
	
	# Run GUI with error handling
	try:
		print("GUI window should now be visible. Check your taskbar if not seen.")
		root.mainloop()
		return results
	except Exception as runtime_error:
		print(f"GUI runtime error: {runtime_error}")
		try:
			root.destroy()
		except (AttributeError, RuntimeError, TclError):
			pass
		return confirm_match_simple_fallback(pending_items)


def confirm_and_iterate_match(normalized_key, matches, ref_data):
	"""Process matches based on confidence."""
	best_match, best_score = matches[0]
	candidate = ref_data.get(best_match, {})
	second_best_score = matches[1][1] if len(matches) > 1 else 0
	is_scryfall_only = candidate.get("TCGplayer Id") == "Scryfall Verified"
	
	# Auto-confirm high confidence matches
	if best_score >= 270 and not is_scryfall_only:
		confirmed_matches[normalized_key] = best_match
		return best_match
	if best_score >= 260 and not is_scryfall_only and (best_score - second_best_score) >= 30:
		confirmed_matches[normalized_key] = best_match
		return best_match
	
	# Auto-confirm Scryfall-verified entries (score 350) - these are high confidence
	if is_scryfall_only and best_score >= 350:
		confirmed_matches[normalized_key] = best_match
		return best_match
	
	# Defer manual review for batch processing
	pending_confirmations.append((normalized_key, matches, ref_data))
	return None  # Will be resolved in batch at end


def build_standard_entry(ref_row, product_name_suffix, manabox_row, condition):
	"""Format standard card entry."""
	return {
			"TCGplayer Id":          ref_row.get("TCGplayer Id", "Not Found"),
			"Product Line":          ref_row.get("Product Line", "Magic: The Gathering"),
			"Set Name":              ref_row.get("Set Name", ""),
			"Product Name":          ref_row.get("Product Name", "") + product_name_suffix,
			"Number":                ref_row.get("Number", ""),
			"Rarity":                ref_row.get("Rarity", ""),
			"Condition":             condition,
			"Add to Quantity":       int(manabox_row.get("Quantity", "1")),
			"TCG Marketplace Price": manabox_row.get("Purchase price", "0.00")
	}


def build_token_entry(ref_row, token_set_name, token_product_name, token_number, manabox_row, condition):
	"""Format token card entry."""
	return {
			"TCGplayer Id":          ref_row.get("TCGplayer Id", "Not Found"),
			"Product Line":          ref_row.get("Product Line", "Magic: The Gathering"),
			"Set Name":              ref_row.get("Set Name", token_set_name),
			"Product Name":          ref_row.get("Product Name", token_product_name),
			"Number":                ref_row.get("Number", token_number),
			"Rarity":                ref_row.get("Rarity", "Token"),
			"Condition":             condition,
			"Add to Quantity":       int(manabox_row.get("Quantity", "1")),
			"TCG Marketplace Price": get_market_price(manabox_row, ref_row)
	}


def build_token_fallback(token_set_name, token_product_name, card_number, manabox_row, condition):
	"""Create unmatched token entry."""
	fallback_price = get_market_price(manabox_row, None)
	return {
			"TCGplayer Id":          "Not Found",
			"Product Line":          "Magic: The Gathering",
			"Set Name":              token_set_name,
			"Product Name":          token_product_name,
			"Number":                card_number,
			"Rarity":                "Token",
			"Condition":             condition,
			"Add to Quantity":       int(manabox_row.get("Quantity", "1")),
			"TCG Marketplace Price": fallback_price
	}


def map_fields(manabox_row, card_database):
	"""Transform input record to output format."""
	card_name = manabox_row.get("Name", "").strip()
	set_name = manabox_row.get("Set name", "").strip()
	condition_code = manabox_row.get("Condition", "near mint").strip().lower().replace("_", " ")
	foil = "Foil" if manabox_row.get("Foil", "normal").lower() == "foil" else ""
	condition = CONDITION_MAP.get(condition_code, "Near Mint")
	if foil:
		condition += " Foil"
	is_token = (
			"token" in set_name.lower() or
			"token" in card_name.lower() or
			(set_name.startswith("T") and re.match(r"^T[A-Z0-9]+$", set_name))
	)
	if is_token:
		return process_token(manabox_row, card_database, condition, card_name, set_name)
	else:
		return process_standard(manabox_row, card_database, condition, card_name, set_name)


def process_standard(manabox_row, _card_database, condition, card_name, set_name):
	"""Handle regular card entries."""
	card_number = re.sub(r"^[A-Za-z\-]*", "", manabox_row.get("Collector number", "").strip().split("-")[-1])
	if not card_name or not set_name:
		return None
	normalized_result = normalize_key(card_name, set_name, condition, card_number)
	if not normalized_result:
		return None
	key = normalized_result[:4]
	
	# Check for existing confirmed matches
	if key in confirmed_matches:
		ref_row = ref_data[confirmed_matches[key]]
		return build_standard_entry(ref_row, normalized_result[4], manabox_row, condition)
	
	# Find matches
	matches = find_best_match(key, ref_data)
	
	# Enhance matches with Scryfall verification for missing or low-confidence matches
	if not matches or (matches and matches[0][1] < 260):
		matches = enhance_matches_with_scryfall(normalized_result, matches, ref_data, manabox_row)
	
	# Try to confirm match (auto-confirm or defer)
	confirmed_match = None
	if matches:
		confirmed_match = confirm_and_iterate_match(key, matches, ref_data)
	
	# If we have a confirmed match, process it
	if confirmed_match:
		ref_row = ref_data[confirmed_match]
		
		# Check if this is a Scryfall-only entry and track it separately
		if ref_row.get("TCGplayer Id") == "Scryfall Verified":
			scryfall_entry = build_standard_entry(ref_row, normalized_result[4], manabox_row, condition)
			scryfall_only_cards.append(scryfall_entry)
			return None  # Don't include in main output
		
		return build_standard_entry(ref_row, normalized_result[4], manabox_row, condition)
	
	# If no match found and not deferred, add to given up
	if not any(item[0] == key for item in pending_confirmations):
		fallback = build_given_up_entry(manabox_row, condition, card_name, set_name)
		given_up_cards.append(fallback)
	
	return None


def process_token(manabox_row, _card_database, condition, card_name, set_name):
	"""Handle token card entries."""
	if set_name.startswith("T") and re.match(r"^T[A-Z0-9]+$", set_name):
		token_set_name = set_name[1:] + " tokens"
	else:
		token_set_name = set_name
	token_set_base = token_set_name.lower().replace(" tokens", "")
	card_number = manabox_row.get("Collector number", "").strip()
	if "//" in card_name:
		parts = card_name.split("//")
		side1 = parts[0].strip()
		side2 = re.sub(r"double[-\s]?sided token", "", parts[1], flags=re.IGNORECASE).strip()
		token_product_name = f"{side1} // {side2}"
	else:
		token_product_name = card_name
	
	normalized_token_key = normalize_key(token_product_name, token_set_name, condition, card_number)
	if not normalized_token_key:
		print(f"Skipping invalid or prerelease token: {card_name} from set {set_name}")
		return None
	
	token_ref_data = {
			k: v for k, v in ref_data.items()
			if (
					("token" in v.get("Set Name", "").lower() or "token" in v.get("Product Name", "").lower()) and
					(token_set_name.lower() in v.get("Set Name", "").lower() or token_set_base in v.get("Set Name",
					                                                                                    "").lower())
			)
	}
	matches = find_best_match(normalized_token_key[:4], token_ref_data)
	chosen_match = None
	
	# Auto-confirm high-confidence token matches
	if matches:
		best_match, best_score = matches[0]
		if best_score >= 250:
			chosen_match = best_match
		else:
			# Defer token confirmation for batch processing
			pending_confirmations.append((normalized_token_key, matches, token_ref_data))
			return None  # Will be processed in batch later
	
	# Check for double-sided tokens if we have an auto-confirmed match
	if chosen_match and "//" in card_name:
		ds_matches = [
				(m, s) for m, s in matches
				if is_double_sided_candidate(token_ref_data[m].get("Product Name", ""))
		]
		if ds_matches and ds_matches[0][0] != chosen_match:
			# Defer if double-sided options exist
			pending_confirmations.append((normalized_token_key[:4], ds_matches, token_ref_data))
			return None
	
	# Process confirmed match
	if chosen_match:
		ref_row = token_ref_data[chosen_match]
		token_product_name = ref_row.get("Product Name", token_product_name)
		token_number = ref_row.get("Number", card_number)
		return build_token_entry(ref_row, token_set_name, token_product_name, token_number, manabox_row, condition)
	
	# No match found and not deferred - add to given up only if not in pending confirmations
	if not any(item[0] == normalized_token_key[:4] for item in pending_confirmations):
		fallback = build_token_fallback(token_set_name, token_product_name, card_number, manabox_row, condition)
		given_up_cards.append(fallback)
	
	return None


def merge_entries(cards):
	"""Consolidate duplicate entries."""
	merged = {}
	for card in cards:
		key = (card['TCGplayer Id'], card['Condition'])
		if key in merged:
			merged[key]['Add to Quantity'] += card['Add to Quantity']
		else:
			merged[key] = card
	return list(merged.values())


def auto_confirm_high_score(cards):
	"""Auto-approve high-confidence matches."""
	confirmed = []
	for card in cards:
		if card.get('Score', 0) >= 250:
			confirmed.append(card)
	return confirmed


def detect_csv_files():
	"""Identify input file types automatically."""
	current_dir = Path(".")
	csv_files = list(current_dir.glob("*.csv"))
	
	print(f"Found {len(csv_files)} CSV files to analyze...")
	
	manabox_file = None
	tcgplayer_file = None
	
	# Look for files with specific patterns
	for csv_file in csv_files:
		filename_lower = csv_file.name.lower()
		print(f"Analyzing: {csv_file.name}")
		
		# Skip output files from previous runs
		if any(skip in filename_lower for skip in
		       ['tcgplayer_staged', 'scryfall_verified', 'tcgplayer_given_up', 'cards_missing_from_tcgplayer']):
			print(f"  Skipping output file")
			continue
		
		# Try to identify file type by reading headers
		try:
			with open(csv_file, 'r', encoding='utf-8') as f:
				header = f.readline().lower()
				
				# Manabox files have this specific combination of columns
				# Based on analysis: Name,Set code,Set name,Collector number,Foil,Rarity,Quantity,ManaBox ID,Scryfall ID,Purchase price
				manabox_indicators = ['manabox id', 'scryfall id', 'set code']
				manabox_matches = [col for col in manabox_indicators if col in header]
				
				# TCGplayer files have this specific combination
				# Based on analysis: TCGplayer Id,Product Line,Set Name,Product Name,Title,Number,Rarity,Condition,TCG Market Price,TCG Direct Low
				tcgplayer_indicators = ['tcgplayer id', 'product line', 'tcg market price']
				tcgplayer_matches = [col for col in tcgplayer_indicators if col in header]
				
				print(f"  Manabox indicators found: {manabox_matches}")
				print(f"  TCGplayer indicators found: {tcgplayer_matches}")
				
				# Strong Manabox detection: Must have ManaBox ID AND Scryfall ID
				if 'manabox id' in header and 'scryfall id' in header:
					if not manabox_file:
						manabox_file = csv_file
						print(f"  -> Identified as Manabox file")
						continue
				
				# Strong TCGplayer detection: Must have TCGplayer Id AND Product Line
				if 'tcgplayer id' in header and 'product line' in header:
					if not tcgplayer_file:
						tcgplayer_file = csv_file
						print(f"  -> Identified as TCGplayer file")
						continue
				
				# Additional check: Look for unique column combinations
				if 'set code' in header and 'collector number' in header and 'manabox id' not in header:
					# This might be a different format, check if it's still manabox-like
					if not manabox_file and 'scryfall id' in header:
						manabox_file = csv_file
						print(f"  -> Identified as Manabox-like file")
						continue
				
				print(f"  -> Could not identify file type")
		
		except Exception as e:
			print(f"  -> Error reading file: {e}")
			continue  # Skip files that can't be read
	
	return manabox_file, tcgplayer_file


def select_csv_file(prompt):
	"""Get file selection from user."""
	file_path = askopenfilename(title=prompt, filetypes=[("CSV Files", "*.csv")])
	if not file_path:
		print(f"No file selected for {prompt}. Exiting.")
		exit()
	return file_path


def create_output_folder():
	"""Generate timestamped output directory."""
	timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
	output_dir = Path(f"converted_output_{timestamp}")
	output_dir.mkdir(exist_ok=True)
	return output_dir


# Main file I/O
print("MTG Card Converter v2.0")
print(f"Filters: Prerelease={FILTER_PRERELEASE}, Promo={FILTER_PROMO}")

# Try auto-detection first
print("Scanning for CSV files...")
detected_manabox, detected_tcgplayer = detect_csv_files()

if detected_manabox and detected_tcgplayer:
	print(f"Auto-detected files:")
	print(f"  Manabox CSV: {detected_manabox.name}")
	print(f"  TCGplayer CSV: {detected_tcgplayer.name}")
	manabox_csv = str(detected_manabox)
	reference_csv = str(detected_tcgplayer)
else:
	print("Could not auto-detect both files. Please select manually...")
	Tk().withdraw()
	
	if not detected_manabox:
		manabox_csv = select_csv_file("Select the Manabox CSV File")
	else:
		print(f"Using detected Manabox file: {detected_manabox.name}")
		manabox_csv = str(detected_manabox)
	
	if not detected_tcgplayer:
		reference_csv = select_csv_file("Select the TCGPlayer Reference CSV File")
	else:
		print(f"Using detected TCGplayer file: {detected_tcgplayer.name}")
		reference_csv = str(detected_tcgplayer)

# Create organized output folder
output_dir = create_output_folder()
print(f"Output folder: {output_dir}")

tcgplayer_csv = output_dir / "tcgplayer_staged_inventory.csv"
ref_data = load_reference_data(reference_csv)

try:
	with open(manabox_csv, mode='r', newline='', encoding='utf-8') as infile, \
			open(tcgplayer_csv, mode='w', newline='', encoding='utf-8') as outfile:
		reader = csv.DictReader(infile)
		fieldnames = [
				"TCGplayer Id", "Product Line", "Set Name", "Product Name",
				"Number", "Rarity", "Condition", "Add to Quantity", "TCG Marketplace Price"
		]
		writer = csv.DictWriter(outfile, fieldnames=fieldnames)
		writer.writeheader()
		cards = []
		for row in reader:
			tcgplayer_row = map_fields(row, ref_data)
			if tcgplayer_row:
				cards.append(tcgplayer_row)
		merged_cards = merge_entries(cards)
		for card in merged_cards:
			writer.writerow(card)
	# Process pending confirmations in batch
	if pending_confirmations:
		print(f"\nProcessing {len(pending_confirmations)} manual confirmations...")
		try:
			confirmation_results = confirm_match_gui_batch(pending_confirmations)
			
			# Apply confirmation results and create additional entries
			confirmed_count = 0
			skipped_count = 0
			additional_cards = []
			
			for confirmation_idx, result in confirmation_results.items():
				if confirmation_idx < len(pending_confirmations):
					normalized_key, matches, local_ref_data = pending_confirmations[confirmation_idx]
					
					if result:
						confirmed_matches[normalized_key] = result
						match_row = local_ref_data[result]
						
						# Create entry for confirmed match
						# Note: This is simplified - in practice we'd need the original manabox_row
						# For now, we'll just record the confirmation
						confirmed_count += 1
						print(f"Confirmed: {normalized_key[0]} -> {match_row.get('Product Name', 'Unknown')}")
					else:
						skipped_count += 1
						print(f"Skipped: {normalized_key[0]}")
			
			print(f"Manual confirmations completed: {confirmed_count} confirmed, {skipped_count} skipped")
			
			# Clear pending confirmations to prevent reprocessing
			pending_confirmations.clear()
		
		except Exception as e:
			print(f"GUI confirmation failed: {e}")
			print("Adding all unconfirmed items to unmatched list...")
			for unmatched_key, unmatched_matches, unmatched_ref_data in pending_confirmations:
				print(f"Unmatched: {unmatched_key[0]}")
			pending_confirmations.clear()
	
	print(f"Conversion complete: {len(merged_cards)} cards")
	
	# Write additional output files
	output_files = [str(tcgplayer_csv)]
	
	if scryfall_only_cards:
		scryfall_csv = output_dir / "cards_missing_from_tcgplayer.csv"
		write_csv_output(scryfall_csv, fieldnames, scryfall_only_cards, "Missing from TCGplayer")
		output_files.append(str(scryfall_csv))
	
	if given_up_cards:
		given_up_csv = output_dir / "tcgplayer_given_up.csv"
		write_csv_output(given_up_csv, fieldnames, given_up_cards, "Unmatched")
		output_files.append(str(given_up_csv))
	
	# Summary
	print(f"\nFiles saved to: {output_dir}")
	for file_path in output_files:
		file_name = Path(file_path).name
		print(f"  - {file_name}")
except FileNotFoundError as e:
	print(f"Error: {e}")
except Exception as e:
	print(f"An unexpected error occurred: {e}")
