from flask import Flask, render_template, request, url_for, redirect, g, flash, jsonify, send_file, session, flash, get_flashed_messages, render_template_string, make_response,send_from_directory,abort
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, current_user, login_required
from flask_paginate import Pagination, get_page_parameter
import pyodbc
import openpyxl
from openpyxl.styles import Font, PatternFill, NamedStyle
import pandas as pd
from pandas import DataFrame, ExcelWriter,to_numeric,to_datetime, to_timedelta
from datetime import datetime, timedelta
import os
from werkzeug.utils import secure_filename
import re
from functools import wraps
import logging
from logging.handlers import RotatingFileHandler
from threading import Thread
import numpy as np
import urllib.parse
import asyncio
import json
from io import BytesIO
import time
import subprocess
from waitress import serve
import sys
from configparser import ConfigParser