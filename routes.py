from io import BytesIO
import bcrypt
import pandas as pd
from flask import Flask, jsonify, redirect, render_template, request, send_file, session, url_for, flash
from flask_login import login_user, current_user, logout_user, login_required
from matplotlib import pyplot as plt
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from app import app, db
from forms import RegistrationForm, LoginForm, UpdateProfileForm, ChangePasswordForm
from models import User
from models import Report

@app.route("/")
@app.route("/home")
def home():
    return render_template('index.html')

@app.route("/upload", methods=['GET', 'POST'])
@login_required
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        graph_type = request.form['graph_type']
        x_axis = request.form.get('x_axis')
        y_axis = request.form.get('y_axis')
        
        if file and file.filename.endswith('.xlsx'):
            df = pd.read_excel(file)
            if graph_type == 'highchart':
                categories = df[x_axis].tolist()
                series_data = df[y_axis].tolist()
                highchart_img = create_highchart_image(categories, series_data, 'line')
                
                # Embed image into Excel file
                wb = openpyxl.load_workbook(file)
                ws = wb.active
                img_data = OpenpyxlImage(highchart_img)
                ws.add_image(img_data, 'E5')
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                return send_file(output, download_name='output.xlsx', as_attachment=True)

            fig, ax = plt.subplots()
            if graph_type == 'pie':
                df.plot.pie(y=x_axis, ax=ax, autopct='%1.1f%%')
            elif graph_type == 'scatter':
                df.plot.scatter(x=x_axis, y=y_axis, ax=ax)
            elif graph_type == 'bar':
                df.plot.bar(x=x_axis, y=y_axis, ax=ax)
            elif graph_type == 'histogram':
                df[x_axis].plot.hist(ax=ax)
            elif graph_type == 'line':
                df.plot.line(x=x_axis, y=y_axis, ax=ax)

            img = BytesIO()
            plt.savefig(img, format='png')
            img.seek(0)
            plt.close()

            # Embed image into Excel file
            wb = openpyxl.load_workbook(file)
            ws = wb.active
            img_data = OpenpyxlImage(img)
            ws.add_image(img_data, 'E5')
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            # Save report to database
            report = Report(title=file.filename, file_path=f'reports/{file.filename}')
            db.session.add(report)
            db.session.commit()

            return send_file(output, download_name='output.xlsx', as_attachment=True)

    return render_template('upload.html')

@app.route("/get_columns", methods=['POST'])
@login_required
def get_columns():
    file = request.files['file']
    if file and file.filename.endswith('.xlsx'):
        df = pd.read_excel(file)
        columns = df.columns.tolist()
        return jsonify(columns)
    return jsonify([])

def mock_login(username):
    session['user'] = username

@app.route('/dashboard')
def dashboard():
    if 'user' in session:
        return render_template('dashboard.html', username=session['user'])
    else:
        return redirect(url_for('login'))

@app.route("/register", methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('home'))
    form = RegistrationForm()
    if form.validate_on_submit():
        user = User(username=form.username.data, email=form.email.data)
        user.set_password(form.password.data)
        db.session.add(user)
        db.session.commit()
        flash('Your account has been created!', 'success')
        return redirect(url_for('login'))
    return render_template('register.html', form=form)

@app.route('/login', methods=['GET', 'POST'])
def login():
    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(email=form.email.data).first()
        if user and user.check_password(form.password.data):
            login_user(user, remember=form.remember.data)
            flash('Login successful!', 'success')
            return redirect(url_for('upload_dashboard'))
        else:
            flash('Invalid email or password. Please try again.', 'danger')
    return render_template('login.html', form=form)

@app.route('/upload/dashboard')
@login_required
def upload_dashboard():
    return render_template('upload.html')

@app.route("/logout")
def logout():
    logout_user()
    return redirect(url_for('home'))

@app.route('/profile', methods=['GET', 'POST'])
@login_required
def profile():
    profile_form = UpdateProfileForm(obj=current_user)
    password_form = ChangePasswordForm()

    if profile_form.validate_on_submit():
        current_user.email = profile_form.email.data
        db.session.commit()
        flash('Your profile has been updated!', 'success')
        return redirect(url_for('profile'))  # Redirect back to profile page

    if password_form.validate_on_submit():
        if bcrypt.check_password_hash(current_user.password, password_form.current_password.data):
            hashed_password = bcrypt.generate_password_hash(password_form.new_password.data).decode('utf-8')
            current_user.password = hashed_password
            db.session.commit()
            flash('Your password has been updated!', 'success')
            return redirect(url_for('profile'))  # Redirect back to profile page
        else:
            flash('Current password is incorrect. Please try again.', 'danger')

    return render_template('profile.html', form=profile_form, password_form=password_form)

@app.route('/reports')
@login_required
def reports():
    reports = Report.query.all()  # Fetch all reports from the database
    return render_template('reports.html', reports=reports)

@app.errorhandler(404)
def not_found_error(error):
    return render_template('404.html'), 404

@app.errorhandler(500)
def internal_error(error):
    return render_template('500.html'), 500

def create_highchart_image(categories, series_data, chart_type='line'):
    import matplotlib.pyplot as plt
    from matplotlib.ticker import MaxNLocator
    from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
    from matplotlib.figure import Figure

    fig = Figure()
    axis = fig.add_subplot(1, 1, 1)
    
    if chart_type == 'column':
        axis.bar(categories, series_data, color='#7cb5ec', edgecolor='black')
    elif chart_type == 'line':
        axis.plot(categories, series_data, marker='o', linestyle='-', color='#7cb5ec')

    axis.set_xlabel('Categories')
    axis.set_ylabel('Values')
    axis.set_title('Highchart Visualization')
    axis.xaxis.set_major_locator(MaxNLocator(integer=True))
    
    # Style to resemble Highcharts
    fig.tight_layout()
    fig.subplots_adjust(top=0.88)
    
    # Improve the appearance
    axis.grid(True, which='both', linestyle='--', linewidth=0.5)
    axis.set_facecolor('#ffffff')
    
    # Hide the top and right spines
    axis.spines['top'].set_visible(False)
    axis.spines['right'].set_visible(False)
    
    # Customize the tick parameters
    axis.tick_params(axis='both', which='major', labelsize=10)
    
    canvas = FigureCanvas(fig)
    img = BytesIO()
    canvas.print_png(img)
    img.seek(0)
    
    return img
