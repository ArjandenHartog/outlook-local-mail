import base64
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace as NS
import unittest
from unittest.mock import patch

import signatures


class SignatureTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name)
        self.files = self.path/'signatures'
        self.files.mkdir()
        self.settings = self.path/'settings.json'
        self.name = 'Jane (a@example.test)'
        (self.files/(self.name+'.htm')).write_text('<div>Kind regards, Jane</div>')
        (self.files/(self.name+'.txt')).write_text('Kind regards, Jane')
        (self.files/'Other (b@example.test).htm').write_text('<div>Other</div>')
        self.settings.write_text(json.dumps({'a@example.test':{'new':self.name,'reply':self.name}}))
        self.patchers = [patch.object(signatures,'SIGNATURES',self.files),patch.object(signatures,'SETTINGS',self.settings)]
        for p in self.patchers:
            p.start()

    def tearDown(self):
        for p in self.patchers:
            p.stop()
        self.temp.cleanup()

    def test_account_scope_and_wrong_account_rejected_before_saving(self):
        self.assertEqual(signatures.select('a@example.test').stem,self.name)
        self.assertEqual(len(signatures.catalog('a@example.test')['signatures']),1)
        with self.assertRaises(ValueError):
            signatures.select('a@example.test',name='Other (b@example.test)')
        with self.assertRaises(ValueError):
            signatures.select('a@example.test',name='../settings')

    def test_template_is_available_but_never_appended_as_signature(self):
        template = self.files/'Acquisitie (a@example.test).htm'
        template.write_text('<p>Offer</p>')
        template.with_suffix('.txt').write_text('A complete outreach message '*30+'Kind regards, Jane')
        self.assertTrue(next(s for s in signatures.catalog('a@example.test')['signatures'] if s['name']==template.stem)['template'])
        with self.assertRaises(ValueError):
            signatures.select('a@example.test',name=template.stem)

    def test_missing_or_explicit_none_signature_stays_none(self):
        self.assertIsNone(signatures.select('empty@example.test'))
        self.assertIsNone(signatures.select('a@example.test',name='none'))

    def test_html_message_is_escaped_and_reply_quote_is_preserved(self):
        draft = NS(HTMLBody='<html><body><p>Original quote</p></body></html>')
        with patch.object(signatures,'attach_images',side_effect=lambda d,c:(c,0)):
            signatures.compose(draft,'Ha,\n<hello>',self.files/(self.name+'.htm'),reply=True)
        self.assertIn('Ha,<br>&lt;hello&gt;',draft.HTMLBody)
        self.assertIn('<p>Original quote</p>',draft.HTMLBody)
        self.assertEqual(draft.HTMLBody.count('Kind regards'),1)

    def test_inline_image_is_embedded_with_matching_cid_and_hidden_property(self):
        records = []
        class Attachments:
            def Add(inner,filename,kind):
                record = {'bytes':Path(filename).read_bytes(),'properties':{}}
                records.append(record)
                return NS(PropertyAccessor=NS(SetProperty=lambda key,val:record['properties'].__setitem__(key,val)))
        encoded = base64.b64encode(b'image bytes').decode()
        draft = NS(Attachments=Attachments())
        (self.path/'runtime').mkdir()
        with patch.object(signatures,'ROOT',self.path):
            content,count = signatures.attach_images(draft,'<img src="data:image/png;base64,'+encoded+'">')
        self.assertEqual(count,1)
        self.assertEqual(records[0]['bytes'],b'image bytes')
        self.assertIn('cid:'+records[0]['properties'][signatures.CID_PROPERTY],content)
        self.assertTrue(records[0]['properties'][signatures.HIDDEN_PROPERTY])

    def test_mailto_link_matches_displayed_account(self):
        path = self.files/'link.htm'
        path.write_text('<!--[if supportFields]>HYPERLINK "mailto:bram@example.test"<![endif]--><a href="mailto:bram@example.test"><b>jane@example.test</b></a>')
        self.assertIn('href="mailto:jane@example.test"',signatures.fragment(path))
        self.assertNotIn('mailto:bram@example.test',signatures.fragment(path))


if __name__ == '__main__':
    unittest.main(verbosity=2)
