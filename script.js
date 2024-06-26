document.addEventListener('DOMContentLoaded', fetchData);

function fetchData() {
    const url = 'E:/Projects/Stacked Bar bot to top file upload/teams.xlsx'; // Update this URL to your hosted Excel file

    fetch(url)
        .then(response => {
            if (!response.ok) {
                throw new Error(`Network response was not ok ${response.statusText}`);
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json(worksheet);
            generateChart(json);
        })
        .catch(error => {
            console.error('Error fetching the Excel file:', error);
        });
}

function generateChart(data) {
    const margin = { top: 60, right: 20, bottom: 100, left: 80 };
    const width = 960 - margin.left - margin.right;
    const height = 600 - margin.top - margin.bottom;

    d3.select("#chart").selectAll("*").remove();

    const svg = d3.select("#chart")
        .append("svg")
        .attr("width", width + margin.left + margin.right)
        .attr("height", height + margin.top + margin.bottom)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);

    const x = d3.scaleBand()
        .domain(data.map(d => d.Team))
        .range([0, width])
        .padding(0.2);

    const y = d3.scaleLinear()
        .domain([0, d3.max(data, d => d.Kills + d.Damage)])
        .nice()
        .range([height, 0]);

    const defs = svg.append("defs");

    const gradient = defs.append("linearGradient")
        .attr("id", "gradient")
        .attr("x1", "0%")
        .attr("x2", "100%")
        .attr("y1", "0%")
        .attr("y2", "0%");
    gradient.append("stop")
        .attr("offset", "0%")
        .style("stop-color", "#555")
        .style("stop-opacity", 1);
    gradient.append("stop")
        .attr("offset", "100%")
        .style("stop-color", "#999")
        .style("stop-opacity", 1);

    const killsGradient = defs.append("linearGradient")
        .attr("id", "kills-gradient")
        .attr("x1", "0%")
        .attr("x2", "0%")
        .attr("y1", "0%")
        .attr("y2", "100%");
    killsGradient.append("stop")
        .attr("offset", "0%")
        .style("stop-color", "#1f77b4")
        .style("stop-opacity", 1);
    killsGradient.append("stop")
        .attr("offset", "100%")
        .style("stop-color", "#5bc0de")
        .style("stop-opacity", 1);

    const damageGradient = defs.append("linearGradient")
        .attr("id", "damage-gradient")
        .attr("x1", "0%")
        .attr("x2", "0%")
        .attr("y1", "0%")
        .attr("y2", "100%");
    damageGradient.append("stop")
        .attr("offset", "0%")
        .style("stop-color", "#ff7f0e")
        .style("stop-opacity", 1);
    damageGradient.append("stop")
        .attr("offset", "100%")
        .style("stop-color", "#f0ad4e")
        .style("stop-opacity", 1);

    svg.append("g")
        .attr("class", "x-axis")
        .attr("transform", `translate(0,${height})`)
        .call(d3.axisBottom(x));

    svg.append("g")
        .attr("class", "y-axis")
        .call(d3.axisLeft(y));

    const barGroups = svg.selectAll(".bar-group")
        .data(data)
        .enter()
        .append("g")
        .attr("class", "bar-group")
        .attr("transform", d => `translate(${x(d.Team)},0)`);

    barGroups.append("rect")
        .attr("class", "bar kills")
        .attr("x", 0)
        .attr("y", height)
        .attr("width", x.bandwidth())
        .attr("height", 0)
        .transition()
        .duration(1000)
        .delay((d, i) => i * 200)
        .attr("y", d => y(d.Kills))
        .attr("height", d => height - y(d.Kills));

    barGroups.append("text")
        .attr("class", "text-inside-bar")
        .attr("x", x.bandwidth() / 2)
        .attr("y", d => y(d.Kills / 2))
        .text(d => d.Kills)
        .attr("opacity", 0)
        .transition()
        .duration(1000)
        .delay((d, i) => i * 200)
        .attr("opacity", 1);

    barGroups.append("rect")
        .attr("class", "bar damage")
        .attr("x", 0)
        .attr("y", height)
        .attr("width", x.bandwidth())
        .attr("height", 0)
        .transition()
        .duration(1000)
        .delay((d, i) => i * 200 + 500)
        .attr("y", d => y(d.Kills + d.Damage))
        .attr("height", d => height - y(d.Damage));

    barGroups.append("text")
        .attr("class", "text-inside-bar")
        .attr("x", x.bandwidth() / 2)
        .attr("y", d => y(d.Kills + d.Damage - d.Damage / 2))
        .text(d => d.Damage)
        .attr("opacity", 0)
        .transition()
        .duration(1000)
        .delay((d, i) => i * 200 + 500)
        .attr("opacity", 1);

    barGroups.append("image")
        .attr("xlink:href", d => `logos/${d.Team}.png`)
        .attr("class", "team-logo image-fly-in")
        .attr("x", -25)
        .attr("y", height + 50)
        .attr("width", 50)
        .attr("height", 50)
        .attr("opacity", 0)
        .transition()
        .duration(1000)
        .delay((d, i) => i * 200 + 1000)
        .attr("y", d => y(d.Kills + d.Damage) + 10)
        .attr("opacity", 1);

    svg.append("text")
        .attr("class", "bar-title")
        .attr("x", width / 2)
        .attr("y", -20)
        .text("Team Comparison");
}
