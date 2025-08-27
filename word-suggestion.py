import win32com.client
import os
from fast_diff_match_patch import diff

INSERTION_TYPE = 1
DELETION_TYPE = 2

OPEN_TAG_LENGTH = 3
CLOSE_TAG_LENGTH = 4

TAGS = set(['<b>', '<u>', '<i>', '<h>', '<\\b>', '<\\u>', '<\\i>', '<\\h>'])



def get_doc(path):
    word = win32com.client.Dispatch("Word.Application")
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")
    doc = word.Documents.Open(path)
    return doc


# removes all suggested deletions, keeps suggested insertions
def get_text(doc):
    visible_text = ""
    start = doc.Content.Start
    end = doc.Content.End
    curr = start

    deletion_ranges = [
        (rev.Range.Start, rev.Range.End)
        for rev in doc.Revisions
        if rev.Type == DELETION_TYPE
    ]

    while curr < end:
        is_deleted = False

        for del_start, del_end in deletion_ranges:
            if del_start <= curr < del_end:
                is_deleted = True
                break

        if not is_deleted:
            range = doc.Range(Start=curr, End=curr + 1)
            visible_text += range.Text.replace('\r', '\n')

        curr += 1
    
    return visible_text


# helpers 
# will likely need to modify accordingly once we figure out llm return structure and tendencies
def clean_llm_text(text):
    text = text.replace('\n', '\r').strip()
    return text


def handle_tags(doc, addition, bold_open, underline_open, italicize_open, highlight_open, curr_index):
    count = 0
    for i in range(len(addition)):
        if i + OPEN_TAG_LENGTH <= len(addition):
            if addition[i: i + OPEN_TAG_LENGTH] in TAGS:
                tag = addition[i: i + OPEN_TAG_LENGTH]
                if tag == '<b>':
                    bold_open.append(curr_index + i - count)
                elif tag == '<u>':
                    underline_open.append(curr_index + i - count)
                if tag == '<i>':
                    italicize_open.append(curr_index + i - count)
                elif tag == '<h>':
                    highlight_open.append(curr_index + i - count)
                count += OPEN_TAG_LENGTH

        if i + CLOSE_TAG_LENGTH <= len(addition):
            if addition[i:i + CLOSE_TAG_LENGTH] in TAGS:
                tag = addition[i:i + CLOSE_TAG_LENGTH]
                if tag == '<\\b>' and bold_open:
                    start = bold_open.pop()
                    end = i + curr_index - count
                    print(start, end)
                    style_range = doc.Range(Start=start, End=end)
                    style_range.Font.Bold = True
                elif tag == '<\\u>' and underline_open:
                    style_range = doc.Range(Start=underline_open.pop(), End=i + curr_index - count)
                    style_range.Font.Underline = 1
                if tag == '<\\i>' and italicize_open:
                    style_range = doc.Range(Start=italicize_open.pop(), End=i + curr_index - count)
                    style_range.Font.Italic  = True
                elif tag == '<\\h>' and highlight_open:
                    style_range = doc.Range(Start=highlight_open.pop(), End=i + curr_index - count)
                    style_range.HighlightColorIndex = 7  # neon yellow
                count += CLOSE_TAG_LENGTH

# main function
def make_suggestions(doc, new_text):
    try:
        original_text = doc.Content.Text.strip()
        new_text = clean_llm_text(new_text)
        changes = diff(original_text, new_text, counts_only=False, as_patch=False)
        print(changes)

        # Review mode 
        doc.TrackRevisions = True

        bold_open = []
        underline_open = []
        italicize_open = []
        highlight_open = []


        index = 0
        for operation, content in changes:
            if operation == "+":
                clean_content = content
                for tag in TAGS:
                    clean_content = clean_content.replace(tag, '')
                

                insert_range = doc.Range(Start=index, End=index)
                insert_range.InsertBefore(clean_content)

                handle_tags(doc=doc, addition=content, bold_open=bold_open, underline_open=underline_open, 
                            italicize_open=italicize_open, highlight_open=highlight_open, curr_index=index)

                index += len(clean_content)

            elif operation == "-":
                replace_range = doc.Range(Start=index, End=index + len(content))
                replace_range.Text = ""

                index += len(content)
            
            elif operation == "=":
                index += len(content)

        # Save
        doc.TrackRevisions = False
        doc.SaveAs(path)

    except Exception as e:
        doc.TrackRevisions = False
        doc.SaveAs(path)
        print(e)


# testing
path = r"C:\Users\zroy1\Downloads\ScripterTest.docx"
doc = get_doc(path)
make_suggestions(doc=doc, new_text="A <b><u>quick<\\u> brown dog<\\b> leaps over the lazy fox.")
get_text(doc=doc)
