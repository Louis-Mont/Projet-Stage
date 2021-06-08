class State:
    """
    Generates a state that can be used in the configuration of tkinter objects
    """

    def __init__(self, stateyes='normal', stateno='disabled'):
        """
        Generates the two component of the state : a enable and disable value
        :param stateyes:
        :param stateno:
        """
        self._yes = stateyes
        self._no = stateno
        self.val = True

    def config(self, c, dflt=True):
        """
        Configure a tkinter object depending on its default value
        :param c: The tkinter object to be configured
        :param dflt: If by default the state is to be enabled or disabled
        """
        c.config(state=self._state(not (self.val ^ dflt)))

    def configs(self, confd_dflt):
        """
        Configures various object that can have a different value
        :param confd_dflt: Dictionary { key : the object to be configured, value : bool default value}
        """
        for c, i in confd_dflt.items():
            self.config(c, i)

    def _state(self, state):
        return self._yes if state else self._no
