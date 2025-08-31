from config import client
from pydantic import BaseModel


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
                        - site_dimension: width and length of the terrain
                        - site_setback: mandatory minimum distance between the building and the terrain boundaries
                        - room_count: number of rooms in the apartment
                        - total_area: total area of the apartment
                        - room_types: types of rooms in the apartment
                        - rooms: list of rooms in the apartment
                        Each room should have the following properties:
                        - type: function of room (e.g. living room, kitchen, bedroom, bathroom), If multiple rooms share the same function, append a number (e.g., "bedroom 1", "bedroom 2")
                        - window_counts: number of windows in the room. If the user does not specify a value, always set window_counts = 1. Do not infer or guess based on the room type (e.g., living room, bedroom, kitchen). Only use the user-provided value or the default (1).
                        - area: area of the room, each room's area should be a realistic number and can include decimals (e.g., 14.3, 17.8). Avoid rounding to multiples of 5 or 10.
                        - connects_to: list of rooms that the current room connects to
                        - sun_hours_summer: quantity of hours a room should receive during summer time; if unknown, write 0
                        - sun_hours_winter: quantity of hours a room should receive during winter time; if unknown, write 0
                        - noise_level: if the room can have low, medium or high level of noise; if the user did not specify, leave this field null; do not invent values
                        
                        Output your response in the following JSON format.  
                        Never leave required fields missing, except those explicitly allowed to be null. 
                        If you cannot determine a value:
                        - For strings: use "none"
                        - For numbers or integers: use 0
                        Do not invent values for user-specified constraints
                        Do not leave any required field missing, except fields that are explicitly allowed to be null.
                        Always ensure all required fields are present and filled according to these rules.                
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
                        "site_setback":{
                            "front":{"type": "number"},
                            "side":{"type": "number"},
                            "rear":{"type": "number"},
                        },
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
                                    "sun_hours_summer": {"type": "integer"},
                                    "sun_hours_winter": {"type": "integer"},
                                    "noise_level": {"type": ["string", "null"]}                               
                                },
                                "required": ["type", "window_counts", "area", "connects_to", "sun_hours_summer", "sun_hours_winter", "noise_level"]
                            }
                        }
                    },
                    "required": ["building_type", "site_location", "site_dimension", "room_count", "total_area", "room_types", "rooms"],
                    "additionalProperties": False
                },
                "strict": True
            }
        }
    }
    result = client.run(**data)
    response_data = result["result"]["response"]
    return response_data

