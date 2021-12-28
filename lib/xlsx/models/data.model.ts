export interface IData {
    type: 'line' | 'bar' | 'pie' | 'scatter';
    title: {
        name: string,
        color?: string,
        size?: number
    };
    range: string,
    rgbColors?: string[],
    labels?: boolean,
    marker?: {
        size?: number;
        shape?: string;
    },
    lineWidth?: number;
}

export interface IPPTChartData {
    type: 'line' | 'bar';
    title: {
        name: string,
        color?: string,
        size?: number
    };
    range?: string,
    data?: any[][] | IPPTChartDataVal[],
    rgbColors?: string[],
    labels?: boolean,
    marker?: {
        size?: number;
        shape?: string;
    },
    lineWidth?: number;
    location?: {
        x?: string;
        y?: string;
        cx?: string;
        cy?: string;
    }
}

export interface IPPTChartDataVal {
    labels: string[];
    name: string,
    values: number[]
}

export interface IPptTableOpt {
    x?: string,
    y?: string,
    cx?: string,
    cy?: string,
    colWidth?: number,
    rowHeight?: number
}