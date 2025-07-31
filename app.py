import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="B2B Grade-wise Projection", layout="wide")
st.title("🎓 B2B Current AY Vs Next AY Projection Dashboard")

grades = ['PN', 'N', 'K1', 'K2', 'G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7', 'G8', 'G9']
growth_grades = ['PN', 'N', 'K1', 'K2', 'G1']

# --- 📄 Upload Format Guide ---
st.markdown("### 📄 Upload Format Guide")
st.markdown("""
Please upload an Excel file in the following format with these **required columns**:

- `Academic Year`
- `School Code`
- `School Name`
- `Zone`
- `Stock Type`
- `Product Type`
- `Ratio`
- All grades: `PN`, `N`, `K1`, `K2`, `G1` to `G9`
""")

# Sample data to show
sample_data = pd.DataFrame({
    'Academic Year': ['AY25-26'] * 4,
    'School Code': ['1100001544', '1100000176', '1100001526', '1100001513'],
    'School Name': ['Orchids Global School', 'SNBP Keshavnagar', 'Priyadarshi UP School', 'United Kids School'],
    'Zone': ['South/East', 'West', 'South/East', 'South/East'],
    'Stock Type': ['AY 24-25', 'AY 25-26', 'AY 25-26', 'AY 25-26'],
    'Product Type': ['Core'] * 4,
    'Ratio': ['01:01'] * 4,
    'PN': [0, 15, 0, 0],
    'N': [20, 80, 25, 20],
    'K1': [25, 65, 27, 20],
    'K2': [25, 80, 38, 20],
    'G1': [25, 170, 65, 20],
    'G2': [25, 170, 46, 20],
    'G3': [25, 160, 36, 20],
    'G4': [25, 154, 35, 20],
    'G5': [25, 135, 36, 20],
    'G6': [10, 130, 0, 0],
    'G7': [0, 135, 0, 0],
    'G8': [0, 113, 0, 0],
    'G9': [0, 0, 0, 0]
})

with st.expander("📊 View Sample Format Table"):
    st.dataframe(sample_data, use_container_width=True)

sample_excel = io.BytesIO()
with pd.ExcelWriter(sample_excel, engine='xlsxwriter') as writer:
    sample_data.to_excel(writer, index=False, sheet_name="Sample_Format")
sample_excel.seek(0)

st.download_button(
    label="📥 Download Sample Excel Format",
    data=sample_excel,
    file_name="AY25-26_Sample_Format.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# --- 📤 File Upload ---
uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()

        required_columns = ['School Code', 'School Name', 'Product Type', 'Stock Type'] + grades
        if 'Stock Type (Book Edition)' in df.columns:
            df.rename(columns={'Stock Type (Book Edition)': 'Stock Type'}, inplace=True)

        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            st.error(f"❌ Missing columns: {missing_columns}")
        else:
            st.success("✅ File uploaded and columns are valid.")
            st.subheader("📘 School Strength (Sample)")
            st.dataframe(df[required_columns].head(), use_container_width=True)

            existing_school_count = df['School Code'].nunique()
            st.info(f"🏫 Existing Schools: {existing_school_count}")

            new_school_count = st.number_input("➕ Expected number of new schools in Next AY", min_value=0, value=100)

            admission_type = st.radio("➕ New Admission in Existing Schools", ["Percentage Growth", "Fixed Number per Grade"])
            if admission_type == "Percentage Growth":
                growth_percentage = st.slider("📈 Growth % in Existing Schools (New Admissions)", 0, 100, 10)
            else:
                fixed_new_admissions = st.number_input("📘 Fixed number of new admissions per grade in each existing school", min_value=0, value=10)

            group_cols = ['Product Type', 'Stock Type']
            strength_ay25_26 = df.groupby(group_cols)[grades].sum().reset_index()

            # --- Promotion Logic ---
            promoted_rows = []
            for _, group_df in df.groupby(group_cols):
                promo = group_df.copy()
                for i in range(len(grades) - 1, 0, -1):
                    promo[grades[i]] = group_df[grades[i - 1]]
                promo['PN'] = 0
                promo['G9'] = 0
                promo['Product Type'] = group_df['Product Type'].iloc[0]
                promo['Stock Type'] = group_df['Stock Type'].iloc[0]
                promoted_rows.append(promo)
            promoted_df = pd.concat(promoted_rows)
            promoted_totals = promoted_df.groupby(group_cols)[grades].sum().reset_index()

            # --- New Admissions ---
            new_admissions = promoted_totals.copy()
            for grade in grades:
                if grade in growth_grades:
                    if admission_type == "Percentage Growth":
                        new_admissions[grade] = (strength_ay25_26[grade] * (growth_percentage / 100)).round().astype(int)
                    else:
                        new_admissions[grade] = fixed_new_admissions * existing_school_count
                else:
                    new_admissions[grade] = 0

            school_counts = df.groupby(group_cols)['School Code'].nunique().reset_index()
            school_counts.rename(columns={'School Code': 'School Count'}, inplace=True)

            avg_per_school = strength_ay25_26.merge(school_counts, on=group_cols)
            for grade in grades:
                avg_per_school[grade] = (avg_per_school[grade] / avg_per_school['School Count']).round()

            new_school_students = avg_per_school.copy()
            for grade in grades:
                new_school_students[grade] = (avg_per_school[grade] * new_school_count).round().astype(int)

            final_totals = promoted_totals.copy()
            for grade in grades:
                final_totals[grade] += new_admissions[grade] + new_school_students[grade]

            # --- Combine for Comparison ---
            comparison_rows = []
            for i in range(len(strength_ay25_26)):
                base = strength_ay25_26.loc[i]
                promo = promoted_totals.loc[i]
                new_exist = new_admissions.loc[i]
                new_sch = new_school_students.loc[i]
                final = final_totals.loc[i]
                for grade in grades:
                    comparison_rows.append({
                        'Product Type': base['Product Type'],
                        'Stock Type': base['Stock Type'],
                        'Grade': grade,
                        'Present Strength': base[grade],
                        'Promoted Strength': promo[grade],
                        'New Admissions (Existing)': new_exist[grade],
                        'New (New Schools)': new_sch[grade],
                        'Total Projected': final[grade]
                    })
            comparison_df = pd.DataFrame(comparison_rows)
            overall = comparison_df.groupby("Grade")[['Present Strength', 'Total Projected']].sum().reset_index()

            # --- Dashboard Tabs ---
            st.subheader("📚 Projection Steps with Logic Explanation")
            tabs = st.tabs([
                "Present Strength",
                "Promoted Strength",
                "New Admissions (Existing)",
                "New School Students",
                "Final Projection",
                "Grade-wise Chart",
            ])

            with tabs[0]:
                st.markdown("### 📘 Present Strength")
                st.markdown("**Formula:** Sum of students per grade grouped by Product Type and Stock Type")
                st.dataframe(strength_ay25_26, use_container_width=True)

            with tabs[1]:
                st.markdown("### 🔁 Promoted Strength")
                st.markdown("Simulated grade promotion: `G{i} = G{i-1}`, `PN = 0`, `G9 = 0`")
                st.dataframe(promoted_totals, use_container_width=True)

            with tabs[2]:
                st.markdown("### ➕ New Admissions (Existing Schools)")
                if admission_type == "Percentage Growth":
                    st.markdown(f"**Growth % applied:** {growth_percentage}% for {growth_grades}")
                else:
                    st.markdown(f"**Fixed admissions per grade per school:** {fixed_new_admissions}")
                st.dataframe(new_admissions, use_container_width=True)

            with tabs[3]:
                st.markdown("### 🏫 New School Students")
                st.markdown(f"**New schools projected:** {new_school_count}")
                st.dataframe(new_school_students, use_container_width=True)

            with tabs[4]:
                st.markdown("### 📊 Final Projection (All Sources)")
                st.markdown("Final = Promoted + New (Existing) + New (New Schools)")
                st.dataframe(comparison_df, use_container_width=True)

            with tabs[5]:
                st.markdown("### 📈 Grade-wise Summary")
                st.bar_chart(overall.set_index("Grade"))

            # --- Excel Export ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                strength_ay25_26.to_excel(writer, index=False, sheet_name='Present Strength')
                school_counts.to_excel(writer, index=False, sheet_name='School Counts')
                avg_per_school.to_excel(writer, index=False, sheet_name='Avg per School')
                new_school_students.to_excel(writer, index=False, sheet_name='New School Students')
                new_admissions.to_excel(writer, index=False, sheet_name='New Admissions (Existing)')
                promoted_totals.to_excel(writer, index=False, sheet_name='Promoted Strength')
                comparison_df.to_excel(writer, index=False, sheet_name='Final Projection')

            output.seek(0)
            st.download_button(
                label="📥 Download Detailed Excel Report",
                data=output,
                file_name="Detailed_Projection_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("👆 Upload your Excel file to begin.")
