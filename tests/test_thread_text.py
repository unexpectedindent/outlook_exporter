import sys
import unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).parents[1]))
from thread_text import split_thread, compare_quote

class ThreadTextTests(unittest.TestCase):
    def test_outlook_header(self):
        raw = 'My answer\n\nFrom: A <a@example.org>\nSent: Tue\nTo: B\nSubject: Hello\n\nOriginal content\nline 2'
        lead, quoted, detected = split_thread(raw)
        self.assertEqual(lead, 'My answer')
        self.assertIn('Original content', quoted)
        self.assertTrue(detected)

    def test_no_header(self):
        self.assertEqual(split_thread('Hello\nworld'), ('Hello\nworld','',False))

    def test_inline_change(self):
        old = 'Decision\nLine one\nBudget is 100\nLine three\nLine four\nLine five'
        new = 'Decision\nLine one\nBudget is 200\nLine three\nLine four\nLine five'
        result = compare_quote(new,old,{'min_quoted_lines':3,'similarity_threshold':0.7,'ignore_signature_from':[]})
        self.assertIn('budget is 200',result)

if __name__ == '__main__':
    unittest.main()
