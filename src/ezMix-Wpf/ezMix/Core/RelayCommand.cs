using System;
using System.Windows.Input;

namespace ezMix.Core
{
    public class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Func<object, bool> _canExecute;

        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter) =>
            _canExecute == null || _canExecute(parameter);

        public void Execute(object parameter) =>
            _execute(parameter);

        public void RaiseCanExecuteChanged() =>
            CommandManager.InvalidateRequerySuggested();
    }

    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;
        private readonly Func<T, bool> _canExecute;

        public RelayCommand(Action<T> execute, Func<T, bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            if (_canExecute == null) return true;

            if (parameter == null)
            {
                // Nếu T là value type thì trả về default(T)
                if (typeof(T).IsValueType)
                    return _canExecute(default);
                return _canExecute((T)(object)null);
            }

            // Check kiểu an toàn
            if (parameter is T value)
                return _canExecute(value);

            return _canExecute(default);
        }

        public void Execute(object parameter)
        {
            T value;
            if (parameter == null)
            {
                value = typeof(T).IsValueType ? default : (T)(object)null;
            }
            else if (parameter is T castValue)
            {
                value = castValue;
            }
            else
            {
                value = default;
            }

            _execute(value);
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public void RaiseCanExecuteChanged() =>
            CommandManager.InvalidateRequerySuggested();
    }
}
