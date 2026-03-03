import re

with open("palimpsest-reader/lib/content.ts", "r", encoding="utf-8") as f:
    text = f.read()

# remove the second abstract
pattern = r',\s*\{\s*"title": "ABSTRACT"[\s\S]*?"keywords": "[^"]*"\s*\}'
text = re.sub(pattern, '', text)

with open("palimpsest-reader/lib/content.ts", "w", encoding="utf-8") as f:
    f.write(text)

with open("palimpsest-reader/components/ArticleReader.tsx", "r", encoding="utf-8") as f:
    tsx = f.read()

tsx = tsx.replace('className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-12 border-b-4 border-black pb-12 print:border-b-2"', 'className="mb-12 border-b-4 border-black pb-12 print:border-b-2"')

prose_old = 'className="prose prose-sm sm:prose-base prose-black max-w-none prose-p:leading-relaxed prose-p:text-black prose-headings:font-sans prose-headings:font-black prose-headings:uppercase prose-blockquote:border-l-4 prose-blockquote:border-black prose-blockquote:bg-gray-100 prose-blockquote:py-4 prose-blockquote:px-6 prose-blockquote:not-italic prose-blockquote:text-black prose-img:w-full prose-img:my-8 prose-a:text-black prose-a:font-bold prose-a:underline prose-a:decoration-2 prose-a:underline-offset-4 hover:prose-a:bg-black hover:prose-a:text-white"'
prose_new = 'className="prose prose-sm sm:prose-base prose-black max-w-none prose-p:leading-relaxed prose-p:text-black prose-headings:font-sans prose-headings:font-black prose-headings:uppercase prose-blockquote:border-l-4 prose-blockquote:border-black prose-blockquote:bg-gray-100 prose-blockquote:py-4 prose-blockquote:px-6 prose-blockquote:not-italic prose-blockquote:text-black prose-img:w-full prose-img:my-8 prose-a:text-black prose-a:font-bold prose-a:underline prose-a:decoration-2 prose-a:underline-offset-4 hover:prose-a:bg-black hover:prose-a:text-white columns-1 md:columns-2 gap-12 print:columns-1"'

tsx = tsx.replace(prose_old, prose_new)

with open("palimpsest-reader/components/ArticleReader.tsx", "w", encoding="utf-8") as f:
    f.write(tsx)
