import re
from bs4 import BeautifulSoup
from markdownify import MarkdownConverter

SHAREPOINT_BASE = "https://gccprod.sharepoint.com"

STRIP_TAGS = [
    "b", "strong", "i", "em", "u", "span", "small",
    "sub", "sup", "mark", "del", "ins",
]

class SharePointConverter(MarkdownConverter):
    """Custom markdownify converter that handles SharePoint HTML quirks."""

    def convert_a(self, el, text, parent_tags):
        if text.strip().lower() in ("back to top", "~back to top"):
            return ""

        href = el.get("href", "") or ""
        if match := re.match(r"^/:([a-z]):/r(.*)", href):
            el["href"] = SHAREPOINT_BASE + match.group(2)
        elif href.startswith("/sites/"):
            el["href"] = SHAREPOINT_BASE + href

        return super().convert_a(el, text, parent_tags)

    def convert_table(self, el, text, parent_tags):
        result = super().convert_table(el, text, parent_tags)
        result = re.sub(r"(\|[\s|]+\|\n)(\|[-| :]+\|\n)", "", result)
        return result

def remove_sharepoint_metadata(soup: BeautifulSoup) -> None:
    for tag in soup.find_all(["script", "style", "img"]):
        tag.decompose()
    for tag in soup.find_all(attrs={"data-sp-webpartdata": True}):
        del tag["data-sp-webpartdata"]
    for tag in soup.find_all(attrs={"data-sp-controldata": True}):
        del tag["data-sp-controldata"]
    for tag in soup.find_all(attrs={"data-sp-prop-name": True}):
        tag.decompose()

def remove_breadcrumb_like_lines(soup: BeautifulSoup) -> None:
    for p in soup.find_all("p"):
        txt = p.get_text(" ", strip=True)
        if txt.count(">") >= 2 and len(txt) <= 250:
            p.decompose()

def clean_html_to_markdown(html_content: str) -> str:
    soup = BeautifulSoup(html_content, "lxml")
    remove_sharepoint_metadata(soup)
    remove_breadcrumb_like_lines(soup)
    converter = SharePointConverter(bullets="-", strip=STRIP_TAGS, wrap=False)
    md_text = converter.convert_soup(soup)

    md_text = re.sub(r"\n{3,}", "\n\n", md_text)
    md_text = md_text.replace("\u200b", "")
    md_text = re.sub(r'"\},\"containsDynamicDataSource[^\s]*', '', md_text)
    md_text = re.sub(r'","encodedImage":"[^\s]*', '', md_text)
    md_text = re.sub(r'Contact Us"\}.*?(?=\n\S|$)', "", md_text, flags=re.DOTALL)
    md_text = re.sub(r'Module Owner"\}.*?(?=\n\S|$)', "", md_text, flags=re.DOTALL)

    return md_text.strip()

# Vercel serverless entry point
def handler(request):
    try:
        html_content = request.get_data(as_text=True)
        if not html_content.strip():
            return {
                "statusCode": 400,
                "body": "Empty HTML content provided."
            }

        markdown = clean_html_to_markdown(html_content)
        return {
            "statusCode": 200,
            "body": markdown,
            "headers": {
                "Content-Type": "text/markdown"
            }
        }

    except Exception as e:
        return {
            "statusCode": 500,
            "body": f"Error processing HTML: {str(e)}"
        }
