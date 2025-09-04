from fast_diff_match_patch import diff
import difflib

original_text = "The quick brown fox jumps over the lazy dog."
new_text = "A quick brown dog leaps over the lazy fox."
diff = difflib.ndiff(original_text.split(" "), new_text.split(" "))
print(list(diff))