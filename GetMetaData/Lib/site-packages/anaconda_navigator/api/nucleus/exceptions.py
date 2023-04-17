# -*- coding: utf-8 -*-

"""Nucleus API related exception"""

from requests import HTTPError

__all__ = ['LoginRequiredException']


class LoginRequiredException(HTTPError):
    """Basic username/password auth required."""
