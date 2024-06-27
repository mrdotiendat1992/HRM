from flask import Flask, render_template, request, url_for, redirect, flash, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, current_user, login_required
from flask_paginate import Pagination, get_page_parameter
import pyodbc
import openpyxl
import pandas as pd
from datetime import datetime, timedelta
import os
from werkzeug.utils import secure_filename
import re
from functools import wraps
from waitress import serve
import logging
from logging.handlers import RotatingFileHandler
from threading import Thread
import numpy as np