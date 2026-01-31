import React from 'react';

const Card = ({ title, children, className = '' }) => {
    return (
        <div className={`card ${className}`}>
            {title && <h2 className="card-title">{title}</h2>}
            {children}
        </div>
    );
};

export default Card;
