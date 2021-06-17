from tkinter import Button, Label
from unittest import TestCase

from user_interface.types.State import State


class TestState(TestCase):
    def __init__(self, *args, **kwargs):
        super(TestState, self).__init__(*args, **kwargs)
        self.state = State()
        self.btn = Button()
        self.label = Label()

    def test_config(self):
        self.assertEqual(self.btn['state'], 'normal')
        self.state.config(self.btn)
        self.assertEqual(self.btn['state'], 'normal')
        self.state.config(self.btn, True)
        self.assertEqual(self.btn['state'], 'normal')
        self.state.config(self.btn, False)
        self.assertEqual(self.btn['state'], 'disabled')

        self.assertEqual(self.label['state'], 'normal')
        self.state.config(self.label, False)
        self.assertEqual(self.label['state'], 'disabled')

    def test_configs(self):
        cis = {
            self.btn: False,
            self.label: True
        }
        self.state.configs(cis)
        self.assertEqual(self.btn['state'], 'disabled')
        self.assertEqual(self.label['state'], 'normal')

        cis[self.btn] = True
        cis[self.label] = False
        self.state.configs(cis)
        self.assertEqual(self.btn['state'], 'normal')
        self.assertEqual(self.label['state'], 'disabled')
