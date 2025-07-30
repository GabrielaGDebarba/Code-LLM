from config import client
from pydantic import BaseModel

class FloorPlan(BaseModel):
    area: float
    room_name: str
    windows_count: int
# não deveria ser type, window_counts, area, connects_to... ?

def make_floorplan(message):
    # Ensure message is clean of any problematic quotes
    message = str(message).strip("'").replace("'", "")
    
    data = {
        "messages": [
            {
                "role": "system",
                "content": """
                        Your task is to create a floor plan with multiple rooms in an apartment.
                        According to the user's request, carefully think of a layout that would suit its request.
                        - building_type: if it is an apartment or a house
                        - site_location: the city where the site is located to get climate information
                        - site_dimension_width: width of the terrain
                        - site_dimension_length: width of the terrain
                        - site_setback: mandatory minimum distance between the building and the terrain boundaries
                        - entrance_orientation: orientation of the entrance door (e.g. north, south, east, west)
                        - room_count: number of rooms in the apartment
                        - total_area: total area of the apartment
                        - room_types: types of rooms in the apartment
                        - rooms: list of rooms in the apartment
                        Each room should have the following properties:
                        - type: function of room (e.g. living room, kitchen, bedroom, bathroom)
                        - window_counts: number of windows in the room
                        - area: area of the room
                        - connects_to: list of rooms that the current room connects to
                        - solar_orientation: orientation of the room relative to the sun (e.g. north, south, east, west)
                        - noise_level: if the room can have low, medium or high level of noise
                        
                        Output your response in the following JSON format.
                        """,
            },
            {
                "role": "user",
                "content": message,
            },
        ],
        "response_format": {
            "type": "json_schema", 
            "json_schema": {
                "name": "floorplan_output",
                "schema": {
                    "type": "object",
                    "properties": { # AQUI QUE PODES ADICIONAR MAIS PROPRIEDADES GERAIS DA CASA
                        "building_type": {"type": "string"},
                        "site_location": {"type": "string"},
                        "site_dimension":{
                            "width":{"type": "number"},
                            "length": {"type": "number"},
                        },
                        "site_dimension_length": {"type": "number"},
                        "site_setback":{
                            "front":{"type": "number"},
                            "side":{"type": "number"},
                            "rear":{"type": "number"},
                        },
                        "entrance_orientation": {"type": "string"},
                        "room_count": {"type": "integer"},
                        "total_area": {"type": "number"},
                        "room_types": {
                            "type": "array",
                            "items": {"type": "string"}
                        },
                        "rooms": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": { # AQUI QUE PODES ADICIONAR MAIS PROPRIEDADES DE CADA ROOM
                                    "type": {"type": "string"},
                                    "window_counts": {"type": "integer"},
                                    "area": {"type": "number"},
                                    "connects_to": {
                                        "type": "array",
                                        "items": {"type": "string"}                            
                                    },
                                    "solar_orientation": {"type": "string"},
                                    "noise_level": {"type": "string"}                                
                                },
                                "required": ["type", "window_counts", "area", "connects_to", "solar_orientation", "noise_level"]
                            }
                        }
                    },
                    "required": ["building_type", "site_location", "site_dimension_width", "site_dimension_length", "entrance_orientation" "room_count", "total_area", "room_types", "rooms"],
                    "additionalProperties": False
                },
                "strict": True
            }
        }
    }
    result = client.run(**data)
    response_data = result["result"]["response"]
    return response_data

def generate_concept(message):
    data = {
        "messages": [
            {
                "role": "system",
                "content": """
                        You are a visionary intern at a leading architecture firm.
                        Your task is to craft a short, poetic, and highly imaginative concept for a building design.
                        Weave the initial information naturally into your idea, letting it inspire creative associations and unexpected imagery.
                        Your concept should feel bold, evocative, and memorable — like the opening lines of a story.
                        Keep your response to a maximum of one paragraph.
                        Avoid generic descriptions; instead, focus on mood, atmosphere, and emotional resonance.
                        """,
            },
            {
                "role": "user",
                "content": f"""
                        What is the concept for this building? 
                        Initial information: {message}
                        """,
            },
        ]
    }
    
    result = client.run(**data)
    return result["result"]["response"]

def extract_attributes(message):
    data = {
        "messages": [
            {
                "role": "system",
                "content": """
                        # Instructions #
                        You are a keyword extraction assistant.
                        Your task is to read a given text and extract relevant keywords according to three categories.
                        """,
            },
            {
                "role": "user",
                "content": f"""
                        # GIVEN TEXT # 
                        {message}
                        """,
            },
        ],
        "response_format": {
            "type": "json_schema",
            "json_schema": {
                "name": "attributes_output",
                "schema": {
                    "type": "object",
                    "properties": {
                        "shape": {"type": "string"},
                        "theme": {"type": "string"},
                        "materials": {"type": "string"}
                    },
                    "required": ["shape", "theme", "materials"],
                    "additionalProperties": False
                },
                "strict": True
            }
        }
    }
    
    result = client.run(**data)
    return result["result"]["response"]

def create_question(message):
    data = {
        "messages": [
            {
                "role": "system",
                "content": """
                        # Instruction #
                        You are a thoughtful research assistant specializing in architecture.
                        Your task is to create an open-ended question based on the given text.
                        Your question should invite an answer that points to references to specific brutalist buildings or notable examples.
                        Imagine the question will be answered using a detailed text about brutalist architecture.
                        The question should feel exploratory and intellectually curious.
                        Output only the question, without any extra text.
                        """,
            },
            {
                "role": "user",
                "content": message,
            },
        ]
    }
    
    result = client.run(**data)
    return result["result"]["response"]
