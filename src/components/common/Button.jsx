import React from 'react';

const Button = ({
    children,
    onClick,
    disabled = false,
    variant = 'primary',
    className = '',
    ...props
}) => {
    const variantClass = variant === 'secondary' ? 'btn-secondary' :
                        variant === 'success' ? 'btn-success' :
                        variant === 'warning' ? 'btn-warning' : '';

    return (
        <button
            className={`btn ${variantClass} ${className}`}
            onClick={onClick}
            disabled={disabled}
            {...props}
        >
            {children}
        </button>
    );
};

export default Button;
