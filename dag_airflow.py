from airflow import DAG
from airflow.operators.python_operator import PythonOperator
from datetime import datetime

# Defina o argumento padrão para a DAG
default_args = {
    'owner': 'seu_nome',
    'start_date': datetime(2023, 11, 7),
    'retries': 1,
}

# Crie uma instância da DAG
dag = DAG(
    'minha_dag',
    default_args=default_args,
    description='Executar script Python do Excel',
    schedule_interval=None,  # Defina o agendamento conforme necessário
    catchup=False,
)

# Defina a função que será executada pela tarefa
def execute_python_script():
    import xlrd
    import pandas as pd

    # Coloque o código Python fornecido aqui
    excel_file_path = "ler_tabela_dinamica.py"
    # Resto do seu código...

# Crie uma tarefa PythonOperator que executará o código
execute_script_task = PythonOperator(
    task_id='executar_script',
    python_callable=execute_python_script,
    dag=dag,
)

# Defina a ordem das tarefas
execute_script_task

if __name__ == "__main__":
    dag.cli()
